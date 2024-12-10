# Visio AI-Assisted Plugin - Instructions and Development Guide

## 1. Project Overview

**Objective:** This Visio plugin empowers users with an AI assistant capable of understanding natural language commands to manipulate and edit Visio diagrams in real-time. The plugin features a user-friendly chat interface, seamless AI integration, and efficient performance through concurrent operations.

**Key Features:**

*   **AI-Powered Chat Interface:** Interact with the AI using natural language within a custom task pane.
*   **Real-time Diagram Editing:**  Instantly see changes reflected on the Visio canvas as the AI processes commands.
*   **Visio Automation:**  Automate the creation, connection, styling, grouping, alignment, and distribution of shapes.
*   **Library Management:** Dynamically load Visio stencils and utilize their shapes based on user instructions.
*   **Concurrent Operations:** Leverage multi-threading for responsive UI and background processing of AI requests.
*   **Status Panel:** Monitor the execution status of commands with a dedicated panel that displays success or failure indicators.
*   **Image Upload and Processing:** Upload images and send them to the AI for processing and relevant actions in the diagram.

## 2. Architecture and Technologies

*   **Platform:** .NET Framework
*   **Language:** C#
*   **Development Environment:** Visual Studio with the .NET desktop development workload.
*   **Visio Interop:** `Microsoft.Office.Interop.Visio` for Visio object model manipulation.
*   **AI Model:** Local LLM (e.g., Llama 3) accessible via a RESTful API (n8n workflows).
*   **AI Communication:** `HttpClient` for asynchronous HTTP requests to the n8n webhook endpoints.
*   **JSON Handling:** `Newtonsoft.Json` for serialization and deserialization of commands and responses.
*   **Concurrency:** `System.Threading.Tasks` for asynchronous operations and multi-threading.
*   **Frontend:** Windows Forms for the custom task pane and chat interface.

## 3. Plugin Structure and Components

### 3.1 ThisAddIn.cs

*   **Entry Point:** The main class for the Visio Add-in.
*   **Ribbon Integration:**  Handles the creation and interaction with the custom Ribbon UI (using `RibbonExtension.cs`).
*   **Initialization:** Sets up the `LibraryManager`, `VisioChatManager`, `AIChatPane`, and `VisioCommandProcessor`.
*   **Webhook Listener:** Starts an `HttpListener` to receive commands from n8n.
*   **Model Loading:**  Fetches the list of available models from the Ollama server via n8n.
*   **Shape Catalog Sending:** Sends the Visio shape catalog to n8n for AI awareness.

### 3.2 LibraryManager.cs

*   **Shape Management:** Loads, organizes, and provides access to shapes from Visio stencils.
*   **Shape Catalog:** Builds a catalog of available shapes (categories and names) for the AI.
*   **Visio Actions:** Implements methods to add, connect, style, group, align, distribute shapes, add text, set color, and retrieve shape properties.

### 3.3 VisioCommandProcessor.cs

*   **Command Execution:**  Interprets JSON commands received from the AI (via n8n) and executes them using the `LibraryManager`.
*   **Command Mapping:**  Maps command types (e.g., "CreateShape", "ConnectShapes") to corresponding `LibraryManager` methods.

### 3.4 AIChatPane.cs

*   **Chat Interface:**  Provides the user interface for interacting with the AI, including:
    *   Chat history display (RichTextBox)
    *   Chat input (TextBox)
    *   Send button
    *   Upload image button
    *   Model selection dropdown (ComboBox)
    *   Command status display (ListView)
    *   Status panel toggle button
*   **Image Handling:** Supports drag-and-drop and file selection for image uploads.
*   **UI Updates:** Updates the chat history and command status list, ensuring thread safety using `Invoke`.

### 3.5 ChatManager.cs

*   **AI Communication:** Handles sending messages and images to the n8n `chat-agent` endpoint.
*   **Model Selection:**  Manages the currently selected AI model.
*   **Response Processing:** Parses AI responses, distinguishes between chat messages and commands, and forwards commands to the `VisioCommandProcessor`.

### 3.6 RibbonExtension.cs

*   Defines the custom Ribbon XML for the plugin, creating a tab with buttons for:
    *   Connecting to the AI server
    *   Refreshing the shape library
    *   Adding a test shape
    *   Displaying the selected model
*   **Note**: The current implementation utilizes a custom task pane (AIChatPane) instead of directly embedding controls in the Ribbon. Consider updating the Ribbon to remove unused controls like the category dropdown if the task pane is the primary interface.

## 4. Setup and Installation

### 4.1 Prerequisites

1. **Visio Installation:** Microsoft Visio (a version compatible with `Microsoft.Office.Interop.Visio`).
2. **Development Environment:** Visual Studio (latest version recommended) with the following workloads:
    *   .NET desktop development
    *   Office/SharePoint development
3. **n8n:**  n8n desktop application running locally.
4. **Ollama:** Ollama server running locally with at least one compatible language model (e.g., Llama 3).
5. **Qdrant (Optional):**  For vector database functionalities (if needed by your n8n workflows), run: `docker run -p 6333:6333 qdrant/qdrant:latest`

### 4.2  n8n Workflow Setup

1. **Import OngoingAgent.json:** Import the provided `OngoingAgent.json` into your n8n instance. This workflow defines the logic for:
    *   Receiving user input (chat messages and images).
    *   Routing requests to either a chat model or an action agent.
    *   Utilizing specialized tools (Color Tool, Shape Tool, Size Tool, Position Tool) to extract parameters from user input.
    *   Forwarding commands to the Visio plugin via the webhook listener (`/visio-command/`).
    *   Sending AI responses back to the Visio plugin for chat history updates.

2. **Configure Ollama Credentials:** In the n8n workflow, configure the "Ollama Chat Model" and "Ollama Model" nodes with your Ollama API credentials (if required).

3. **Activate the Workflow:** Ensure the `OngoingAgent` workflow is activated in n8n.

### 4.3 Visio Plugin Setup

1. **Clone the Repository:** Clone your project repository to your local machine.
2. **Open in Visual Studio:** Open the solution (`.sln` file) in Visual Studio.
3. **Restore NuGet Packages:** Build the solution to restore the required NuGet packages (e.g., `Microsoft.Office.Interop.Visio`, `Newtonsoft.Json`).
4. **Configure API Endpoint:** In `ThisAddIn.cs`, verify that the `apiEndpoint` variable is set to your n8n webhook URL (default: `http://localhost:5678/webhook`).
5. **Build and Run:**
    *   Build the solution (Build -> Build Solution).
    *   Start debugging (Debug -> Start Debugging) or press F5. This will launch Visio and load the plugin.

### 4.4 Using the Plugin

1. **Open Visio:** Launch Visio after building the plugin.
2. **AI Chat Pane:** You should see the "AI Chat Pane" docked on the right side of the Visio window. If not, it might be hidden behind other panels.
3. **Connect to n8n:**
    *   In the plugin's Ribbon tab, click the "Connect" button. This will:
        *   Load the available AI models from your Ollama server via n8n.
        *   Send the Visio shape catalog to n8n, making the AI aware of your available shapes.
    *   Select an AI model from the dropdown in the "AI Chat Pane."
4. **Start Chatting:**
    *   Type commands or questions in the chat input box.
    *   Press Enter or click "Send" to send the message to the AI.
5. **Upload Images:**
    *   Click "Upload" to select an image file or drag and drop an image onto the chat history area.
    *   The image will be sent to n8n for processing by the AI.
6. **Observe Real-time Actions:** The AI will process your commands and execute actions in Visio. You'll see the diagram updated in real-time.
7. **Monitor Status:** The "Status" panel (toggleable via the "Status" button) shows the status (Success/Failed) of each executed command.

## 5. Development Guide

### 5.1 Adding New Visio Commands

1. **LibraryManager:** Implement a new public method in `LibraryManager.cs` for the desired Visio action. Use the Visio Interop API (`Microsoft.Office.Interop.Visio`) to interact with the Visio document.
2. **VisioCommandProcessor:**
    *   Add a new `case` to the `ProcessCommand` method's `switch` statement, corresponding to the new command type.
    *   Extract parameters from the `JObject` as needed.
    *   Call the appropriate `LibraryManager` method to execute the action.
3. **n8n Workflow (OngoingAgent.json):**
    *   Create new tools (if necessary) to extract parameters from user input.
    *   Modify the "action_agent" prompt to instruct the AI on how to generate the new command and use the associated tools.
    *   Construct the JSON payload for the new command in the "action_agent" node.
4. **Testing:** Thoroughly test the new command by sending appropriate messages through the chat interface.

### 5.2 Modifying AI Behavior

1. **n8n Workflow:** Adjust the prompts in the "Chat LLM Chain" and "action_agent" nodes to refine the AI's responses and command generation.
2. **Tools:** Modify the code within the tools (Color Tool, Shape Tool, etc.) to change how parameters are extracted or to add new functionalities.

### 5.3 UI Enhancements

1. **AIChatPane:** Modify the `AIChatPane.cs` to add new UI elements, change the layout, or improve the chat interface's behavior.
2. **RibbonExtension:** If needed, update the `RibbonExtension.cs` and the associated XML to add or modify Ribbon controls.

### 5.4 Debugging

1. **Visual Studio Debugger:** Use the Visual Studio debugger to step through the code, inspect variables, and identify issues.
2. **Debug Output:** Utilize `Debug.WriteLine()` statements to log important events and values to the Output window in Visual Studio.
3. **n8n Workflow Editor:** Use the n8n workflow editor to trace the execution flow, inspect data passed between nodes, and debug the AI's logic.

## 6. Troubleshooting

*   **Plugin Not Loading:**
    *   Ensure the plugin is enabled in Visio (File -> Options -> Add-ins).
    *   Check the "Disabled Items" list in Visio.
    *   Verify that the build output path is correct and that the VSTO manifest is properly configured.
*   **Connection Issues:**
    *   Confirm that n8n is running and the `OngoingAgent` workflow is activated.
    *   Check the `apiEndpoint` in `ThisAddIn.cs`.
    *   Verify that your firewall is not blocking the connection.
*   **AI Not Responding:**
    *   Ensure that the Ollama server is running and the selected model is loaded.
    *   Check the n8n workflow logs for errors.
*   **Command Execution Failures:**
    *   Use the Visual Studio debugger to step through the `VisioCommandProcessor` and `LibraryManager` code.
    *   Examine the command status in the "Status" panel.
    *   Check the Visio document for any error messages or unexpected behavior.

## 7. Further Development

*   **Enhanced Error Handling:** Implement more robust error handling throughout the plugin, including better user feedback and error reporting.
*   **Undo/Redo Functionality:**  Integrate AI actions with Visio's undo/redo stack for a more seamless user experience.
*   **Context Awareness:** Enhance the AI's context awareness by providing it with more information about the current state of the Visio diagram (e.g., selected shapes, page properties).
*   **Voice Control:** Explore the possibility of adding voice control to the plugin using speech recognition libraries.
*   **Customizable UI:** Allow users to customize the appearance and layout of the chat interface.
*   **Advanced AI Features:** Integrate more advanced AI features, such as:
    *   **Diagram Generation from Text Descriptions:** Generate entire diagrams from high-level text descriptions.
    *   **Smart Suggestions:** Provide intelligent suggestions to the user based on the current diagram context.
    *   **Automated Diagram Optimization:** Automatically improve the layout and organization of diagrams.