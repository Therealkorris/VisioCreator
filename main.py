from fastapi import FastAPI, WebSocket, HTTPException, Form, WebSocketDisconnect
from fastapi.responses import StreamingResponse
import requests
import json
import logging
from agent import VisioAgent

logging.basicConfig(level=logging.INFO)
app = FastAPI()

visio_agent = VisioAgent()

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


@app.post("/text-prompt")
async def handle_text_prompt(prompt: str = Form(...), model: str = Form("llama3.2")):
    try:
        logging.info(f"Received text prompt: {prompt} for model: {model}")
        
        # Send request to AI API
        payload = {"model": model, "prompt": prompt}
        response = requests.post(f"http://localhost:11434/api/generate", json=payload, stream=True)
        response.raise_for_status()

        # Read and process each chunk of the streamed response as separate JSON objects
        full_response = ""
        for chunk in response.iter_lines():
            if chunk:
                decoded_chunk = json.loads(chunk.decode('utf-8'))
                # Append only the 'response' field to the final response
                if 'response' in decoded_chunk:
                    full_response += decoded_chunk['response'] + " "

        # Log and return the clean AI response
        clean_response = full_response.strip()
        logging.info(f"AI Response Processed: {clean_response}")
        return {"response": clean_response}

    except requests.RequestException as e:
        logging.error(f"Error communicating with AI: {str(e)}")
        return {"response": f"Error communicating with AI: {str(e)}"}


    except requests.RequestException as e:
        logging.error(f"Error communicating with AI: {str(e)}")
        return {"response": f"Error communicating with AI: {str(e)}"}




async def ai_response_stream(prompt: str, model: str):
    try:
        logging.info(f"Sending to AI API - Prompt: {prompt}, Model: {model}")
        response = requests.post(f"http://localhost:11434/api/generate", 
                                 json={"model": model, "prompt": prompt}, 
                                 stream=False)  # Set stream to False to get full response
        response.raise_for_status()

        # Process the response JSON
        ai_response = response.json()

        # Extract only the 'response' field and return it
        full_response = ""
        for chunk in ai_response:
            if 'response' in chunk:
                full_response += chunk['response']

        logging.info(f"AI Response Received: {full_response.strip()}")
        return full_response.strip()

    except requests.RequestException as e:
        logging.error(f"Error communicating with AI: {str(e)}")
        return f"Error communicating with AI: {str(e)}"


# Agent-based processing of commands for Visio interaction
@app.post("/agent-prompt")
async def handle_agent_prompt(prompt: str = Form(...), model: str = Form("llama3.2")):
    logging.info(f"Received agent prompt: {prompt} for model: {model}")
    
    # Fetch AI response stream
    ai_responses = []
    async for response in ai_response_stream(prompt, model):
        ai_responses.append(response)

    # Join the responses to form a single command text
    command_text = ''.join(ai_responses)
    logging.info(f"AI Command received: {command_text}")

    # Use the agent to parse and execute the command
    command_type, command_data = visio_agent.parse_command(command_text)
    if command_type and command_data:
        agent_response = visio_agent.execute_command(command_type, command_data)
        return {"agent_response": agent_response}
    else:
        return {"error": "Failed to parse or execute command."}

# This function processes AI-generated commands and executes Visio actions
def process_visio_agent_command(command):
    if "create_shape" in command:
        shape = command.get('shape', 'default_shape')
        x = command.get('x', 1.0)
        y = command.get('y', 1.0)
        logging.info(f"Creating shape: {shape} at ({x}, {y}) in Visio.")
        return f"Shape '{shape}' created at coordinates ({x}, {y})."
    return "Command not recognized."

# WebSocket for Visio-specific commands
@app.websocket("/ws/visio-command")
async def websocket_visio_command(websocket: WebSocket):
    await websocket.accept()
    while True:
        try:
            data = await websocket.receive_text()
            logging.info(f"Received Visio command: {data}")
            processed_data = process_visio_agent_command(data)
            await websocket.send_text(f"Processed command: {processed_data}")
        except WebSocketDisconnect:
            logging.info("WebSocket disconnected")
            break

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
