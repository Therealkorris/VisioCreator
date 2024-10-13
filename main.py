from fastapi import FastAPI, HTTPException, Form, WebSocket, WebSocketDisconnect
import logging
from agent import handle_prompt_from_agent, process_visio_agent_command, list_models

logging.basicConfig(level=logging.DEBUG)  # DEBUG level for detailed logs
app = FastAPI()

@app.get("/healthcheck")
async def healthcheck():
    logging.debug("Health check endpoint hit")
    return {"status": "running"}

@app.get("/models")
async def get_models():
    try:
        models = list_models()  # Directly list models using the AI API endpoint
        return {"models": models}
    except Exception as e:
        logging.error(f"Error fetching models: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error fetching models: {str(e)}")

@app.post("/agent-prompt")
async def handle_agent_prompt(prompt: str = Form(...), model: str = Form("llama3.2")):
    try:
        ai_response = await handle_prompt_from_agent(prompt, model)
        return {"response": ai_response}
    except Exception as e:
        logging.error(f"Error processing AI prompt: {str(e)}")
        return {"error": f"Error processing AI prompt: {str(e)}"}

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