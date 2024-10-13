import uuid
import logging
from qdrant_client import QdrantClient
from qdrant_client.http.models import PointStruct  # Import PointStruct

# Initialize Qdrant Client
def initialize_qdrant_client():
    client = QdrantClient(host="localhost", port=6333)
    return client

# Create collections with vector configurations
def create_collections(client):
    try:
        # Create 'models' collection
        client.recreate_collection(
            collection_name="models",
            vectors_config={"size": 1536, "distance": "Cosine"}  # Adjust vector size for model embeddings
        )
        # Create 'actions' collection
        client.recreate_collection(
            collection_name="actions",
            vectors_config={"size": 512, "distance": "Cosine"}  # Adjust vector size for action embeddings
        )
        # Create 'shapes' collection
        client.recreate_collection(
            collection_name="shapes",
            vectors_config={"size": 128, "distance": "Cosine"}  # Adjust vector size for shape data
        )
        # Create 'function_blocks' collection
        client.recreate_collection(
            collection_name="function_blocks",
            vectors_config={"size": 512, "distance": "Cosine"}  # Adjust vector size for function blocks
        )
        logging.info("Collections created successfully in Qdrant.")
    except Exception as e:
        logging.error(f"Error creating collections in Qdrant: {e}")

# Store model in Qdrant with descriptive metadata
def store_model_in_qdrant(client, model_name, model_data):
    try:
        model_id = str(uuid.uuid4())  # Use a UUID for point ID
        client.upsert(
            collection_name="models",  # Separate collection for models
            points=[
                PointStruct(
                    id=model_id,  # Use UUID as ID
                    vector=model_data,  # Model data as vector (embedding)
                    payload={"model_name": model_name}  # Payload with descriptive name
                )
            ]
        )
        logging.info(f"Model {model_name} stored successfully with UUID {model_id}")
    except Exception as e:
        logging.error(f"Error storing model in Qdrant: {e}")

# Store action in Qdrant with descriptive metadata
def store_action_in_qdrant(client, action_name, action_type, action_data):
    try:
        action_id = str(uuid.uuid4())  # Use a UUID for point ID
        client.upsert(
            collection_name="actions",  # Separate collection for actions
            points=[
                PointStruct(
                    id=action_id,  # Use UUID as ID
                    vector=action_data,  # Action data as vector (embedding)
                    payload={
                        "action_name": action_name,
                        "action_type": action_type
                    }
                )
            ]
        )
        logging.info(f"Action {action_name} stored successfully with UUID {action_id}")
    except Exception as e:
        logging.error(f"Error storing action in Qdrant: {e}")

# Store shape data in Qdrant with descriptive metadata
def store_shape_in_qdrant(client, shape_name, shape_data):
    try:
        shape_id = str(uuid.uuid4())  # Use a UUID for point ID
        client.upsert(
            collection_name="shapes",  # Separate collection for shapes
            points=[
                PointStruct(
                    id=shape_id,  # Use UUID as ID
                    vector=shape_data,  # Shape data as vector (embedding)
                    payload={"shape_name": shape_name}
                )
            ]
        )
        logging.info(f"Shape {shape_name} stored successfully with UUID {shape_id}")
    except Exception as e:
        logging.error(f"Error storing shape in Qdrant: {e}")

# Store function block data in Qdrant with descriptive metadata
def store_function_block_in_qdrant(client, block_name, block_data):
    try:
        block_id = str(uuid.uuid4())  # Use a UUID for point ID
        client.upsert(
            collection_name="function_blocks",  # Separate collection for function blocks
            points=[
                PointStruct(
                    id=block_id,  # Use UUID as ID
                    vector=block_data,  # Function block data as vector (embedding)
                    payload={"block_name": block_name}
                )
            ]
        )
        logging.info(f"Function block {block_name} stored successfully with UUID {block_id}")
    except Exception as e:
        logging.error(f"Error storing function block in Qdrant: {e}")

# Fetch data from Qdrant knowledge base
def fetch_data_from_qdrant(client, collection_name, query_vector):
    try:
        result = client.search(
            collection_name=collection_name,
            query_vector=query_vector,  # The query vector (embedding)
            limit=5
        )
        return result
    except Exception as e:
        logging.error(f"Error fetching data from Qdrant: {e}")
        return None

# Function to search for similar actions
def search_similar_actions(client, query_vector, limit=5):
    return fetch_data_from_qdrant(client, "actions", query_vector)
