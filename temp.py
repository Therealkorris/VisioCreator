import logging
import json
import requests
import os
from langchain.schema import Document
from langchain_ollama import ChatOllama
from langchain_community.document_loaders import WebBaseLoader
from langchain_community.vectorstores import SKLearnVectorStore
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_nomic.embeddings import NomicEmbeddings
from typing import Dict, List, Union

logging.basicConfig(level=logging.INFO)

# Initialize LLM
local_llm = "llama3.2:3b-instruct-fp16"
llm = ChatOllama(model=local_llm, temperature=0)
llm_json_mode = ChatOllama(model=local_llm, temperature=0, format="json")

# Load Documents from URLs (Placeholder for future RAG integration)
urls = [
    "https://lilianweng.github.io/posts/2023-06-23-agent/",
    "https://lilianweng.github.io/posts/2023-03-15-prompt-engineering/",
    "https://lilianweng.github.io/posts/2023-10-25-adv-attack-llm/",
]

try:
    loaded_docs = []
    for url in urls:
        loader = WebBaseLoader(url)
        docs = loader.load()
        loaded_docs.extend(docs)
    docs_list = [item for sublist in loaded_docs for item in sublist]
    # Split documents for VectorDB
    text_splitter = RecursiveCharacterTextSplitter.from_tiktoken_encoder(chunk_size=1000, chunk_overlap=200)
    doc_splits = text_splitter.split_documents(docs_list)
    # Add to vectorDB
    vectorstore = SKLearnVectorStore.from_documents(
        documents=doc_splits,
        embedding=NomicEmbeddings(model="nomic-embed-text-v1.5", inference_mode="local"),
    )
    retriever = vectorstore.as_retriever(k=3)
    HARD_CODED_DOCUMENT = loaded_docs[0].page_content if loaded_docs else "Default hardcoded content for testing."
except Exception as e:
    logging.error(f"Failed to load documents from URLs: {e}")
    HARD_CODED_DOCUMENT = "Default hardcoded content for testing."

# Action Functions

def create_shape(command_data: dict) -> str:
    shape = command_data.get("shape")
    x = command_data.get("x")
    y = command_data.get("y")
    width = command_data.get("width", 50)
    height = command_data.get("height", 50)
    color = command_data.get("color", "default")
    return f"Created shape '{shape}' at ({x}, {y}) with width {width} and height {height} in color {color}."

def connect_shapes(command_data: dict) -> str:
    shape1 = command_data.get("shape1")
    shape2 = command_data.get("shape2")
    return f"Connected shape '{shape1}' to shape '{shape2}'."

def modify_properties(command_data: dict) -> str:
    shape = command_data.get("shape")
    property_name = command_data.get("property")
    value = command_data.get("value")
    return f"Modified property '{property_name}' of shape '{shape}' to value '{value}'."

# Ollama Integration

def call_ollama(prompt: str, model: str = "llama3.2") -> str:
    response = llm.invoke([Document(page_content=prompt, metadata={"source": "loaded_document"})])
    return response.content.strip()

# VisioAgent Class

class VisioAgent:
    def __init__(self):
        self.commands = {
            "create_shape": create_shape,
            "connect_shapes": connect_shapes,
            "modify_properties": modify_properties,
        }

    def parse_user_message(self, user_message: str) -> Union[Dict, List[Dict]]:
        prompt = f'''
You are an AI assistant that interprets user requests to perform actions on a canvas.

Given the user's message, extract the actions to be performed and represent them as a JSON array if there are multiple actions, or a JSON object for a single action.

Each action should follow this schema:
[
  {{"action": "create_shape", "shape": "circle", "x": 10, "y": 90, "width": 50, "height": 50, "color": "red"}},
  {{"action": "create_shape", "shape": "circle", "x": 90, "y": 90, "width": 50, "height": 50, "color": "blue"}}
]
Only respond with the JSON object/array, without any additional text.

Do not include any additional text other than the JSON object/array.

User message: "{user_message}"

AI Response:
'''
        full_response = call_ollama(prompt, model=local_llm)
        logging.info(f"VisioAgent AI Response: {full_response}")
        try:
            command_data = json.loads(full_response)
            return command_data
        except json.JSONDecodeError:
            logging.error(f"Failed to parse AI response as JSON: {full_response}")
            return {"error": "Invalid JSON response from AI"}

    def execute_action(self, command_data: Union[Dict, List[Dict]]) -> Union[str, List[str]]:
        if isinstance(command_data, list):
            return [self._execute_single_action(cmd) for cmd in command_data]
        else:
            return self._execute_single_action(command_data)

    def _execute_single_action(self, command_data: Dict) -> str:
        action = command_data.get("action")
        if action in self.commands:
            return self.commands[action](command_data)
        else:
            return f"Unsupported command '{action}'"

# ManagerAgent Function

def manager_agent(user_message: str) -> str:
    routing_prompt = '''
You are an expert at determining the appropriate handling of user requests.

Based on the user's message, decide whether the request should be routed to the VisioAgent for creating or modifying shapes, or handled as a general conversational response, or if additional documents should be retrieved from the vector store.

The VisioAgent should be used for commands that involve creating shapes, connecting shapes, modifying properties, or other canvas-related actions. Use the conversational assistant for all other questions or general requests.

Return JSON with a single key, "route", which can be "visio_agent", "conversational", or "retrieval".

User message: "{user_message}"

AI Response:
'''
    routing_response = call_ollama(routing_prompt.format(user_message=user_message), model=local_llm)
    logging.info(f"Routing Decision: {routing_response}")
    try:
        routing_data = json.loads(routing_response)
        route = routing_data.get("route")
        if route == "visio_agent":
            visio_agent = VisioAgent()
            command_data = visio_agent.parse_user_message(user_message)
            if "error" in command_data:
                return command_data["error"]
            execution_result = visio_agent.execute_action(command_data)
            return json.dumps(execution_result, indent=2)
        elif route == "retrieval":
            retrieved_docs = retriever.invoke(user_message)
            context = "\n\n".join(doc.page_content for doc in retrieved_docs)
            rag_prompt = f"""
You are an assistant for question-answering tasks.

Here is the context to use to answer the question:

{context}

Think carefully about the above context.

Now, review the user question:

{user_message}

Provide an answer to this question using only the above context. Use three sentences maximum and keep the answer concise.

Answer:
"""
            generation = call_ollama(rag_prompt)
            return generation
        else:
            return handle_conversational_response(user_message)
    except json.JSONDecodeError:
        logging.error(f"Failed to parse routing response as JSON: {routing_response}")
        return "Error: Invalid routing response from AI"

# Handle Conversational Response

def handle_conversational_response(user_message: str, model: str = local_llm) -> str:
    prompt = f'''
You are a friendly and helpful assistant.

User: "{user_message}"

Assistant:
'''
    return call_ollama(prompt, model=model)

# Test Cases
def test_cases():
    test_inputs = [
        "Create a blue circle at position (10, 20) with radius 15.",
        "Connect the red square to the right top corner",
        "Modify the color of the blue circle to green.",
        "What are the types of agent memory?",
        "What's the weather like today?",
    ]
    for test_input in test_inputs:
        print(f"User: {test_input}")
        response = manager_agent(test_input)
        print(f"Agent: {response}\n")

if __name__ == "__main__":
    test_cases()