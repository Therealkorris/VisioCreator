import logging
import requests
import json
from tools import create_shape, connect_shapes, modify_shape_properties
from qdrant_db import initialize_qdrant_client, store_model_in_qdrant, store_action_in_qdrant, fetch_data_from_qdrant, search_similar_actions


logging.basicConfig(level=logging.INFO)

class VisioAgent:
    def __init__(self):
        self.commands = {
            "create_shape": self.create_shape,
            "connect_shapes": self.connect_shapes,
            "modify_properties": self.modify_properties,
        }
        self.qdrant_client = initialize_qdrant_client()

    # Ollama embedding generation
    def generate_ollama_embedding(self, text: str, model: str = "mxbai-embed-large"):
        url = "http://localhost:11434/api/embed"
        headers = {"Content-Type": "application/json"}
        
        payload = {
            "model": model,
            "prompt": text
        }

        response = requests.post(url, headers=headers, json=payload)
        
        if response.status_code == 200:
            embedding_data = response.json()
            return embedding_data.get("embedding", [])
        else:
            raise Exception(f"Failed to generate embeddings: {response.status_code}, {response.text}")

    def parse_command(self, command_text):
        logging.info(f"VisioAgent: Parsing command.")
        try:
            logging.info(f"Raw AI response: {command_text}")
            command_data = json.loads(command_text)
            return command_data.get('action'), command_data
        except json.JSONDecodeError as e:
            logging.error(f"Failed to parse JSON: {command_text}, Error: {str(e)}")
            return None, None

    def execute_command(self, command_type, command_data):
        if command_type in self.commands:
            return self.commands[command_type](command_data)
        else:
            return {"error": f"Unsupported command '{command_type}'"}

    def create_shape(self, command_data):
        shape = command_data.get('shape')
        x = command_data.get('x')
        y = command_data.get('y')
        color = command_data.get('color', 'default')

        # Handle radius for circles
        if 'radius' in command_data:
            radius = command_data.get('radius')
            result = create_shape(shape, x, y, radius * 2, radius * 2, color)
        else:
            width = command_data.get('width')
            height = command_data.get('height')
            result = create_shape(shape, x, y, width, height, color)

        # Generate embedding for action and store it
        action_embedding = self.generate_ollama_embedding(json.dumps(command_data))
        store_action_in_qdrant(self.qdrant_client, f"{shape}_{x}_{y}", "create_shape", action_embedding)  # Ensure to pass 'action_embedding'
        return result

    def connect_shapes(self, command_data):
        shape1 = command_data.get('shape1')
        shape2 = command_data.get('shape2')
        result = connect_shapes(shape1, shape2)

        action_embedding = self.generate_ollama_embedding(json.dumps(command_data))
        store_action_in_qdrant(self.qdrant_client, f"connect_{shape1}_{shape2}", "connect_shapes", action_embedding)  # Pass the embedding
        return result

    def modify_properties(self, command_data):
        shape = command_data.get('shape')
        property_name = command_data.get('property')
        value = command_data.get('value')
        result = modify_shape_properties(shape, property_name, value)

        action_embedding = self.generate_ollama_embedding(json.dumps(command_data))
        store_action_in_qdrant(self.qdrant_client, f"modify_{shape}_{property_name}", "modify_properties", action_embedding)  # Pass the embedding
        return result

    def search_similar_actions(self, query_text, limit=5):
        embedding = self.generate_ollama_embedding(query_text)
        return search_similar_actions(self.qdrant_client, embedding, limit)

# Function to handle prompt requests to the AI API
async def handle_prompt_from_agent(prompt: str, model: str = "llama3.2"):
    full_prompt = f'''
Canvas information:
- Size: 100x100 units
- Coordinate system: (0,0) is top-left, (100,100) is bottom-right

User request: {prompt}

Interpret the user's request and provide a response as a JSON array if there are multiple actions, 
or a JSON object for a single action. For example:
[
  {{"action": "create_shape", "shape": "circle", "x": 10, "y": 90, "width": 50, "height": 50, "color": "red"}},
  {{"action": "create_shape", "shape": "circle", "x": 90, "y": 90, "width": 50, "height": 50, "color": "blue"}}
]
Only respond with the JSON object/array, without any additional text.
'''

    try:
        response = requests.post("http://localhost:11434/api/generate", json={"model": model, "prompt": full_prompt}, stream=True)
        response.raise_for_status()

        full_response = ""
        for line in response.iter_lines():
            if line:
                chunk = json.loads(line.decode('utf-8'))
                full_response += chunk.get("response", "")
                if chunk.get("done", False):
                    break

        logging.info(f"Full AI Response: {full_response}")

        # Store the AI response in Qdrant with embeddings
        embedding = VisioAgent().generate_ollama_embedding(full_response)
        store_action_in_qdrant(initialize_qdrant_client(), f"ai_response_{prompt[:20]}", "ai_response", embedding)  # Ensure to pass the embedding

        try:
            command_data = json.loads(full_response)
            return command_data
        except json.JSONDecodeError:
            logging.error(f"Failed to parse AI response as JSON: {full_response}")
            return {"error": "Invalid JSON response from AI"}

    except requests.RequestException as e:
        logging.error(f"Error communicating with AI: {str(e)}")
        return {"error": f"Error communicating with AI: {str(e)}"}

# New function for model listing
def list_models():
    """ Fetches available models from the AI backend """
    try:
        response = requests.get("http://localhost:11434/api/tags")
        response.raise_for_status()
        models_data = response.json()["models"]
        model_names = [model["name"] for model in models_data]
        logging.info(f"Available models: {model_names}")
        return model_names
    except requests.RequestException as e:
        logging.error(f"Error fetching models: {str(e)}")
        raise Exception(f"Error fetching models: {str(e)}")

def process_visio_agent_command(command):
    try:
        command_data = json.loads(command)
        
        if isinstance(command_data, list):
            results = []
            visio_agent = VisioAgent()
            for cmd in command_data:
                action = cmd.get('action')
                if action in visio_agent.commands:
                    result = visio_agent.execute_command(action, cmd)
                    results.append(result)
                else:
                    results.append({"error": f"Unsupported command '{action}'"})
            return results
        else:
            action = command_data.get('action')
            if action:
                return VisioAgent().execute_command(action, command_data)
            else:
                return {"error": "Unrecognized command"}
    except Exception as e:
        logging.error(f"Error processing Visio command: {str(e)}")
        return {"error": f"Error processing command: {str(e)}"}

# New function to search for similar actions
def search_similar_actions(query_text, limit=5):
    visio_agent = VisioAgent()
    return visio_agent.search_similar_actions(query_text, limit)
