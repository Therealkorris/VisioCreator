**Project Overview:**

The Visio AI-Assistant Plugin is a sophisticated tool that bridges the gap between natural language and Visio diagram manipulation. It leverages AI to empower users to control Visio through a chat interface, making diagram creation and editing more intuitive. The core components are:

*   **Visio Plugin (C#):**  Handles user interaction within Visio, communicates with n8n, and executes commands on the Visio document.
*   **n8n Workflow (JSON):**  Acts as the intermediary between Visio and the Ollama AI server. It routes requests, processes responses, and generates JSON commands for Visio.
*   **Ollama Server:** Hosts the AI language models that power the natural language processing and command generation.

**File Breakdown:**

1. **AIChatPane.cs:**
    *   Defines the user interface for the chat pane within Visio.
    *   Handles user input (text and image uploads).
    *   Displays chat history and command status.
    *   Manages model selection.
    *   Sends messages and images to the `VisioChatManager`.
    *   Updates the command status list view.

2. **ChatManager.cs:**
    *   Responsible for sending messages and images to the n8n workflow.
    *   Processes responses from n8n, determining if they are chat messages or Visio commands.
    *   Forwards Visio commands to `VisioCommandProcessor`.
    *   Updates the chat history in `AIChatPane`.

3. **LibraryManager.cs:**
    *   Manages the Visio shape library (stencils).
    *   Provides methods to add, connect, style, group, align, and distribute shapes.
    *   Retrieves shape properties and page size information.
    *   Sends the shape catalog to n8n on startup.

4. **VisioAutomation.cs:** (Currently not extensively used in the main workflow but provides useful functionalities.)
    *   Offers lower-level methods for interacting with Visio documents, stencils, and shapes.
    *   Can be used to extend the plugin's capabilities.

5. **VisioCommandProcessor.cs:**
    *   Parses JSON commands received from n8n.
    *   Executes the corresponding Visio actions using `LibraryManager`.
    *   Handles various command types (e.g., `CreateShape`, `ConnectShapes`, `AddTextToShape`).

6. **ThisAddIn.cs:**
    *   The entry point for the Visio plugin.
    *   Initializes components (`LibraryManager`, `VisioCommandProcessor`, `VisioChatManager`).
    *   Handles Ribbon UI events (e.g., connect, refresh libraries).
    *   Manages the webhook listener for receiving commands from n8n.
    *   Loads available AI models from Ollama (via n8n).

7. **StructuredoutputOllama.json:**
    *   Defines the n8n workflow.
    *   Sets up webhook triggers for receiving messages from Visio.
    *   Uses Ollama nodes for interacting with the AI models.
    *   Includes code nodes for parsing and processing data.
    *   Uses a switch node to route requests based on whether they are chat messages or Visio commands.
    *   Sends HTTP requests to the Visio plugin to execute commands.