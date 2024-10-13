from fastapi import FastAPI, HTTPException, Form, WebSocket, WebSocketDisconnect
import requests
import json
import logging
from agent import VisioAgent

logging.basicConfig(level=logging.DEBUG)  # DEBUG level for detailed logs
app = FastAPI()

visio_agent = VisioAgent()

@app.get("/healthcheck")
async def healthcheck():
    logging.debug("Health check endpoint hit")
    return {"status": "running"}

@app.get("/models")
async def list_models():
    try:
        logging.debug("Fetching available models from AI backend")
        response = requests.get("http://localhost:11434/api/tags")
        response.raise_for_status()
        models_data = response.json()["models"]
        model_names = [model["name"] for model in models_data]
        logging.info(f"Available models: {model_names}")
        return {"models": model_names}
    except requests.RequestException as e:
        logging.error(f"Error fetching models: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error fetching models: {str(e)}")

@app.post("/test-visio-command")
async def test_visio_command():
    try:
        logging.debug("Generating a test command for Visio")
        test_command = {
            "action": "create_shape",
            "shape": "Circle",
            "x": 200,
            "y": 200,
            "width": 50,
            "height": 50,
            "color": "blue"
        }
        logging.info(f"Test Command: {test_command}")

        return {"status": "success", "command": test_command}
    except Exception as e:
        logging.error(f"Error in generating test command: {str(e)}")
        raise HTTPException(status_code=500, detail="Error in generating test command")

@app.post("/agent-prompt")
async def handle_agent_prompt(prompt: str = Form(...), model: str = Form("llama3.2")):
    try:
        logging.debug(f"Processing AI prompt: {prompt} for model: {model}")
        
        ai_responses = await ai_response_stream(prompt, model)
        command_text = ai_responses.strip()

        logging.info(f"AI Command received: {command_text}")
        command_type, command_data = visio_agent.parse_command(command_text)

        if command_type and command_data:
            logging.info(f"Executing command {command_type} with data {command_data}")
            agent_response = visio_agent.execute_command(command_type, command_data)
            return {"agent_response": agent_response}
        else:
            logging.warning("Failed to parse or execute AI command")
            return {"error": "Failed to parse or execute command."}
    except Exception as e:
        logging.error(f"Error processing AI prompt: {str(e)}")
        return {"error": f"Error processing AI prompt: {str(e)}"}

async def ai_response_stream(prompt: str, model: str):
    try:
        logging.info(f"Sending to AI API - Prompt: {prompt}, Model: {model}")
        response = requests.post(f"http://localhost:11434/api/generate", 
                                 json={"model": model, "prompt": prompt}, 
                                 stream=False)
        response.raise_for_status()

        ai_response = response.json()

        full_response = ai_response.get("response", "")

        logging.info(f"AI Response Received: {full_response.strip()}")
        return full_response.strip()

    except requests.RequestException as e:
        logging.error(f"Error communicating with AI: {str(e)}")
        return f"Error communicating with AI: {str(e)}"

# Define process_visio_agent_command to handle Visio commands
def process_visio_agent_command(command):
    try:
        # Example command structure - assuming the command is JSON formatted
        command_data = json.loads(command)

        # Determine the command type and pass it to the VisioAgent for execution
        if 'create_shape' in command_data:
            return visio_agent.execute_command('create_shape', command_data['create_shape'])
        elif 'connect_shapes' in command_data:
            return visio_agent.execute_command('connect_shapes', command_data['connect_shapes'])
        elif 'modify_properties' in command_data:
            return visio_agent.execute_command('modify_properties', command_data['modify_properties'])
        else:
            return {"error": "Unrecognized command"}
    except Exception as e:
        logging.error(f"Error processing Visio command: {str(e)}")
        return {"error": f"Error processing command: {str(e)}"}

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
