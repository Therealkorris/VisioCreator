from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import requests
from typing import List, Dict
import logging

logging.basicConfig(level=logging.INFO)

app = FastAPI()

class Query(BaseModel):
    prompt: str
    model: str = "llama3.2"

class Conversation(BaseModel):
    id: str
    messages: List[Dict[str, str]] = []

conversations: Dict[str, Conversation] = {}

@app.get("/healthcheck")
async def healthcheck():
    return {"status": "running"}

@app.get("/models")
async def list_models():
    try:
        response = requests.get("http://localhost:11434/api/tags")
        response.raise_for_status()
        models_data = response.json()["models"]
        model_names = [model["name"] for model in models_data]
        print(f"Available models: {model_names}")  # Debugging info
        return {"models": model_names}
    except requests.RequestException as e:
        raise HTTPException(status_code=500, detail=f"Error fetching models: {str(e)}")

@app.post("/execute-command")
async def execute_command(request_data: dict):
    prompt = request_data.get("prompt")
    model = request_data.get("model", "llama3:latest")
    
    if not prompt:
        return {"error": "Prompt is missing"}

    try:
        # Sending prompt to Ollama AI API
        response = requests.post(f"http://localhost:11434/api/generate", json={"model": model, "prompt": prompt})
        response.raise_for_status()

        ai_result = response.json()
        return {"response": ai_result.get("response", "No response from AI")}
    
    except requests.RequestException as e:
        return {"error": f"Failed to communicate with AI: {str(e)}"}

# Additional route for handling Visio-specific commands
@app.post("/handle-visio-command")
async def handle_visio_command(request_data: dict):
    # Logic for handling Visio actions
    command = request_data.get("command")
    # process Visio commands like create_shape etc.
    return {"status": "Processed Visio command"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
