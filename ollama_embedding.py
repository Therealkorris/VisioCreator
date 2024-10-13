import requests
import json

def generate_ollama_embedding(text: str, model: str = "llama2"):
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
