
***Visio AI-Assistant Plugin: Overview and Capabilities***

## 1. Introduction

This document provides a comprehensive overview of the Visio AI-Assistant Plugin, a powerful tool designed to enhance your Visio diagramming experience with the help of artificial intelligence. This plugin allows you to interact with an AI assistant directly within Visio, enabling you to create, modify, and manage your diagrams using natural language commands.

## 2. Key Features

### 2.1. AI-Powered Chat Interface

*   **Natural Language Interaction:** Communicate with the AI using everyday language. No need to remember complex commands or menus.
*   **Custom Task Pane:** A dedicated "AI Chat Pane" is integrated into the Visio interface, providing a seamless chat experience.
*   **Chat History:**  The chat pane displays a history of your interactions with the AI, allowing you to easily track your progress.
*   **Model Selection:** Choose from a list of available AI models to find the one that best suits your needs.
*   **Status Panel:** Monitor the execution of your commands with a dedicated panel, displaying success or failure notifications for each action.
*   **Image Upload and Processing:** Upload images directly into the chat and let the AI process them, performing actions based on the image content.

### 2.2. Real-time Diagram Editing

*   **Instant Feedback:** See your changes reflected on the Visio canvas immediately as the AI processes your commands.
*   **Automated Actions:** The AI can perform a wide range of actions, including:
    *   **Adding Shapes:** Create new shapes based on your descriptions (e.g., "add a red circle", "create a square").
    *   **Connecting Shapes:** Automatically connect shapes with appropriate connectors.
    *   **Styling Shapes:** Modify the appearance of shapes, including color, line style, and fill.
    *   **Adding Text:** Add text labels to shapes.
    *   **Grouping and Ungrouping:**  Combine shapes into groups or separate grouped shapes.
    *   **Aligning and Distributing:**  Organize shapes neatly using alignment and distribution commands.
    *   **Retrieving Shape Properties:** Get information about shapes, such as their position, size, color, and text.
    *   **Retrieving Page Size:** Get the dimensions of the current Visio page.
*   **Visual Confirmation:** The status panel provides clear visual feedback on the success or failure of each command.

### 2.3. Intelligent Automation

*   **Library Management:** The plugin can access and utilize shapes from your Visio stencils.
*   **Shape Catalog:** The available shapes are organized into a catalog, making it easy for the AI to understand your requests.
*   **AI-Driven Actions:** The AI interprets your natural language commands and translates them into specific actions within Visio.

### 2.4. Seamless Integration

*   **Visio Add-in:** The plugin is built as a Visio COM Add-in, ensuring deep integration with the Visio application.
*   **n8n Workflow:** The plugin communicates with a powerful n8n workflow that handles the AI processing.
*   **Ollama Server:**  The n8n workflow interacts with a local Ollama server, which hosts the AI language models.

## 3. How It Works

1. **User Input:** You type a command or question into the chat input box in the "AI Chat Pane."
2. **Image Upload (Optional):** You can upload an image by clicking the "Upload" button or dragging and dropping it onto the chat history.
3. **Send to n8n:** The plugin sends your message or image to the `chat-agent` endpoint of your local n8n workflow.
4. **AI Processing (n8n):** The n8n workflow routes your request to the appropriate AI agent (either a chat model or an action agent).
    *   **Specialized Tools:** The action agent utilizes tools (Color Tool, Shape Tool, Size Tool, Position Tool) to extract relevant parameters from your input (e.g., shape type, color, size, position).
    *   **Command Generation:** The action agent constructs a JSON command based on your request and the extracted parameters.
5. **Command Execution (Visio):** The n8n workflow sends the JSON command to the Visio plugin via a webhook listener (`/visio-command/`). The plugin then:
    *   **Interprets the Command:**  The `VisioCommandProcessor` parses the JSON command.
    *   **Executes the Action:** The `LibraryManager` performs the corresponding action in Visio (e.g., adding a shape, connecting shapes).
6. **Feedback and Updates:**
    *   **Chat History:** The AI's response is displayed in the chat history.
    *   **Status Panel:** The status of the command (success or failure) is shown in the status panel.
    *   **Visio Canvas:** The Visio diagram is updated in real-time to reflect the changes.

## 4. Getting Started

1. **Prerequisites:** Ensure you have Visio, Visual Studio (with .NET and Office development workloads), n8n, and Ollama installed and running.
2. **Install the Plugin:** Build the Visio plugin solution in Visual Studio and run it. This will install the plugin into Visio.
3. **Import n8n Workflow:** Import the `OngoingAgent.json` workflow into your n8n instance and activate it.
4. **Connect:** In the plugin's Ribbon tab, click "Connect" to establish communication with the AI server (via n8n).
5. **Select a Model:** Choose an AI model from the dropdown menu in the "AI Chat Pane."
6. **Start Chatting:** Type your commands into the chat input box and press Enter or click "Send."

## 5. Example Commands

Here are a few examples of commands you can use:

*   "Add a blue circle in the center."
*   "Create a rectangle and a square."
*   "Connect the circle and the square."
*   "Make the rectangle red."
*   "Add the text 'Start' to the circle."
*   "Group the circle and the square."
*   "Align the shapes to the left."
*   "Distribute the shapes horizontally."
*   "What is the color of the circle?"
*   "What is the size of the page?"
*   "Upload an image of a flowchart and describe it." (After uploading an image)

## 6. Using this document as an AI Prompt

You can copy and paste the entire content of this document into an AI chatbot or prompt to give it context about your Visio plugin project. This will help the AI better understand your questions and requests related to the plugin's functionality and development.