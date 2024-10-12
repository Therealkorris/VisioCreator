from fastapi import FastAPI, File, UploadFile, WebSocket, HTTPException, Form, WebSocketDisconnect
from fastapi.responses import StreamingResponse
import requests
import json
import logging
import asyncio

logging.basicConfig(level=logging.INFO)

app = FastAPI()

# Health check endpoint
@app.get("/healthcheck")
async def healthcheck():
    return {"status": "running"}

# List models from AI backend
@app.get("/models")
async def list_models():
    try:
        response = requests.get("http://localhost:11434/api/tags")
        response.raise_for_status()
        models_data = response.json()["models"]
        model_names = [model["name"] for model in models_data]
        logging.info(f"Available models: {model_names}")
        return {"models": model_names}
    except requests.RequestException as e:
        logging.error(f"Error fetching models: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error fetching models: {str(e)}")

# AI response streaming generator for text input
async def ai_response_stream(prompt: str, model: str):
    try:
        logging.info(f"Sending to AI API - Prompt: {prompt}, Model: {model}")
        response = requests.post(f"http://localhost:11434/api/generate", json={"model": model, "prompt": prompt}, stream=True)
        response.raise_for_status()

        # Process the response as a stream of JSON fragments
        for chunk in response.iter_lines():
            if chunk:
                decoded_chunk = chunk.decode('utf-8')
                try:
                    # Parse each fragment of JSON
                    ai_result = json.loads(decoded_chunk)
                    content = ai_result.get('response', '')
                    if content:
                        logging.info(f"AI Response Chunk: {content}")
                        yield f"{content}\n"
                except json.JSONDecodeError as e:
                    logging.error(f"Error parsing chunk: {e}")
                    continue
    except requests.RequestException as e:
        logging.error(f"Error communicating with AI: {str(e)}")
        yield f"Error communicating with AI: {str(e)}"

# Endpoint for handling text-only prompts
@app.post("/text-prompt")
async def handle_text_prompt(prompt: str = Form(...), model: str = Form("llama3.2")):
    logging.info(f"Received text prompt: {prompt} for model: {model}")
    return StreamingResponse(ai_response_stream(prompt, model), media_type="text/plain")

# Endpoint for handling image + text multimodal prompts
@app.post("/image-prompt")
async def handle_image_prompt(prompt: str = Form(...), file: UploadFile = File(...), model: str = Form("llama3.2")):
    logging.info(f"Received image prompt: {prompt}, model: {model}, file: {file.filename}")

    # Prepare the files for the request
    files = {"file": (file.filename, await file.read(), file.content_type)}
    
    # Prepare form data
    data = {"prompt": prompt, "model": model}
    
    # Send the request to the AI API
    try:
        response = requests.post("http://localhost:11434/api/generate", data=data, files=files)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        logging.error(f"Error processing the image-prompt: {str(e)}")
        raise HTTPException(status_code=500, detail="Failed to process the image")

# WebSocket for Visio-specific commands
@app.websocket("/ws/visio-command")
async def websocket_visio_command(websocket: WebSocket):
    """Handles Visio-specific commands through WebSocket."""
    await websocket.accept()
    while True:
        try:
            data = await websocket.receive_text()
            logging.info(f"Received Visio command: {data}")
            # Placeholder for Visio command handling
            await websocket.send_text(f"Processed command: {data}")
        except WebSocketDisconnect:
            logging.info("WebSocket disconnected")
            break

# Additional route for handling Visio-specific commands over HTTP
@app.post("/handle-visio-command")
async def handle_visio_command(request_data: dict):
    command = request_data.get("command")
    logging.info(f"Received Visio command: {command}")
    # Process Visio commands like create_shape etc.
    return {"status": "Processed Visio command"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
