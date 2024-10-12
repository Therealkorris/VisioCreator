from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from fastapi.responses import StreamingResponse
import requests
from typing import List, Dict
import logging
import json
import asyncio

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


# AI response streaming generator
async def ai_response_stream(prompt: str, model: str):
    try:
        response = requests.post(f"http://localhost:11434/api/generate", json={"model": model, "prompt": prompt}, stream=True)

        # Process the response as a stream of JSON fragments
        for chunk in response.iter_lines():
            if chunk:
                decoded_chunk = chunk.decode('utf-8')
                try:
                    # Parse each fragment of JSON
                    ai_result = json.loads(decoded_chunk)
                    content = ai_result.get('response', '')
                    if content:
                        yield f"{content}\n"
                except json.JSONDecodeError as e:
                    logging.error(f"Error parsing chunk: {e}")
                    continue

    except requests.RequestException as e:
        yield f"Error communicating with AI: {str(e)}"

# Main chat endpoint with streaming response
@app.post("/execute-command")
async def execute_command(request_data: dict):
    # Log the raw request data for debugging
    print(f"Received request data: {request_data}")

    prompt = request_data.get("prompt")
    model = request_data.get("model", "llama3:latest")
    
    if not prompt:
        return {"error": "Prompt is missing"}

    # Stream the AI response back to the client
    return StreamingResponse(ai_response_stream(prompt, model), media_type="text/plain")


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
