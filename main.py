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
async def execute_command(query: Query):
    try:
        # Sending command to the API server
        response = requests.post(
            "http://localhost:11434/api/generate",
            json={"model": query.model, "prompt": query.prompt}
        )
        response.raise_for_status()
        return {"status": "success", "result": response.json()["response"]}
    except requests.RequestException as e:
        raise HTTPException(status_code=500, detail=f"Error generating response: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
