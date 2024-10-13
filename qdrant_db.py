import uuid
import logging
from qdrant_client import QdrantClient
from qdrant_client.http.models import PointStruct, Distance, VectorParams

# Initialize Qdrant Client
def initialize_qdrant_client():
    client = QdrantClient(host="localhost", port=6333)
    return client

# Function to ensure collections exist or create them if they don't
def ensure_collection_exists(client, collection_name, vector_size):
    try:
        collections = client.get_collections().collections
        if collection_name not in [col.name for col in collections]:
            client.create_collection(
                collection_name=collection_name,
                vectors_config=VectorParams(size=vector_size, distance=Distance.COSINE)
            )
            logging.info(f"Collection '{collection_name}' created with vector size {vector_size}.")
        else:
            logging.info(f"Collection '{collection_name}' already exists.")
    except Exception as e:
        logging.error(f"Error ensuring collection exists: {e}")

# Create all necessary collections
def create_all_collections(client):
    ensure_collection_exists(client, "models", vector_size=1536)
    ensure_collection_exists(client, "actions", vector_size=512)
    ensure_collection_exists(client, "shapes", vector_size=128)
    ensure_collection_exists(client, "function_blocks", vector_size=512)
    logging.info("All collections ensured to exist.")

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
            collection_name="actions",  # Ensure 'actions' collection exists
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

