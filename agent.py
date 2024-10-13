import logging
import requests
import json
from tools import create_shape, connect_shapes, modify_shape_properties

logging.basicConfig(level=logging.INFO)

class VisioAgent:
    def __init__(self):
        self.commands = {
            "create_shape": self.create_shape,
            "connect_shapes": self.connect_shapes,
            "modify_properties": self.modify_properties,
        }

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
        width = command_data.get('width')
        height = command_data.get('height')
        color = command_data.get('color', 'default')
        return create_shape(shape, x, y, width, height, color)

    def connect_shapes(self, command_data):
        shape1 = command_data.get('shape1')
        shape2 = command_data.get('shape2')
        return connect_shapes(shape1, shape2)

    def modify_properties(self, command_data):
        shape = command_data.get('shape')
        property_name = command_data.get('property')
        value = command_data.get('value')
        return modify_shape_properties(shape, property_name, value)


async def handle_prompt_from_agent(prompt: str, model: str = "llama3.2"):
    full_prompt = f'''
Canvas information:
- Size: 100x100 units
- Coordinate system: (0,0) is top-left, (100,100) is bottom-right

User request: {prompt}

Interpret the user's request and provide a response as a JSON object with the following fields:
- 'action': The action to perform (e.g., 'create_shape')
- 'shape': The type of shape to create
- 'x': X-coordinate (0-100)
- 'y': Y-coordinate (0-100)
- 'width': Width of the shape (1-100)
- 'height': Height of the shape (1-100)
- 'color': Color of the shape

Respond ONLY with the JSON object, no additional text.
'''

    try:
        response = requests.post("http://localhost:11434/api/generate", json={"model": model, "prompt": full_prompt}, stream=True)
        response.raise_for_status()

        # Gather all chunks from streaming response
        full_response = ""
        for line in response.iter_lines():
            if line:
                chunk = json.loads(line)
                full_response += chunk.get("response", "")
                if chunk.get("done", False):
                    break

        logging.info(f"Full AI Response: {full_response}")

        # Try to parse the full response as JSON
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
        action = command_data.get('action')

        if action == 'create_shape':
            return VisioAgent().execute_command('create_shape', command_data)
        elif action == 'connect_shapes':
            return VisioAgent().execute_command('connect_shapes', command_data)
        elif action == 'modify_properties':
            return VisioAgent().execute_command('modify_properties', command_data)
        else:
            return {"error": "Unrecognized command"}
    except Exception as e:
        logging.error(f"Error processing Visio command: {str(e)}")
        return {"error": f"Error processing command: {str(e)}"}
