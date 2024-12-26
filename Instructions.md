## 1. Introduction
***Visio AI-Assistant Plugin: Overview and Capabilities***

This document provides a comprehensive overview of the Visio AI-Assistant Plugin, a powerful tool designed to enhance your Visio diagramming experience with the help of artificial intelligence. This plugin allows you to interact with an AI assistant directly within Visio, enabling you to create, modify, and manage your diagrams using natural language commands.

## 2. Key Features

### 2.1. AI-Powered Chat Interface

*   **Natural Language Interaction:** Communicate with the AI using everyday language to create and modify diagrams.
*   **Custom Task Pane:** A dedicated "AI Chat Pane" integrated into the Visio interface.
*   **Chat History:** View and track your conversation history with the AI.
*   **Model Selection:** Choose from available Ollama AI models.
*   **Status Panel:** Monitor command execution with success/failure notifications.
*   **Image Upload:** Upload images for AI processing and diagram creation.

### 2.2. Real-time Diagram Editing

*   **Shape Management:**
    *   **Adding Shapes:** Create shapes with specific types, positions, sizes, and colors.
    *   **Connecting Shapes:** Create straight or curved connectors between shapes.
    *   **Text Labels:** Add and modify text within shapes.
    *   **Styling:** Customize shape appearance with colors, line styles, and fill patterns.
    *   **Grouping:** Combine multiple shapes or ungroup existing shape groups.
    *   **Alignment:** Align shapes horizontally or vertically.
    *   **Distribution:** Distribute shapes evenly across space.
*   **Information Retrieval:**
    *   **Shape Properties:** Get details about shape position, size, and styling.
    *   **Page Dimensions:** Retrieve current page size information.

### 2.3. Intelligent Automation

*   **Stencil Integration:** Access and utilize shapes from installed Visio stencils.
*   **Shape Catalog:** Organized catalog of available shapes for AI reference.
*   **Command Processing:** Automatic translation of natural language to Visio actions.

### 2.4. System Architecture

*   **Visio Add-in:** COM Add-in with deep Visio integration.
*   **n8n Workflow:** Handles AI processing and command generation.
*   **Ollama Integration:** Local AI model hosting and processing.

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
3. **Import n8n Workflow:** Import the `Working__Agent_multi_creation.json`, `Image_agent.json`,`Get_Stensils.json`,`Visio_connection_Ollama.json`, workflow into your n8n instance and activate it.
4. **Connect:** In the plugin's Ribbon tab, click "Connect" to establish communication with the AI server (via n8n).
5. **Select a Model:** Choose an AI model from the dropdown menu in the "AI Chat Pane."
6. **Start Chatting:** Type your commands into the chat input box and press Enter or click "Send."