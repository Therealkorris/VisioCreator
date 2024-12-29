**Project Overview:**

The Visio AI-Assistant Plugin is a sophisticated tool that bridges the gap between natural language and Visio diagram manipulation. It leverages AI to empower users to control Visio through a chat interface, making diagram creation and editing more intuitive. The core components are:

*   **Visio Plugin (C#):** Handles user interaction within Visio, processes commands, and manages shape operations.
*   **n8n Workflow:** Routes messages between the plugin and Ollama, processes AI responses, and generates Visio commands.
*   **Ollama Server:** Provides local AI model hosting for natural language processing.

**File Breakdown:**

1. **AIChatPane.cs:**
    *   Implements the custom task pane UI in Visio.
    *   Manages chat input/output and image uploads.
    *   Displays chat history and command status.
    *   Handles model selection from available Ollama models.
    *   Routes messages to `VisioChatManager`.

2. **ChatManager.cs:**
    *   Manages communication with n8n workflow.
    *   Processes AI responses and command execution.
    *   Routes commands to `VisioCommandProcessor`.
    *   Updates chat history and status.
    *   Handles image upload processing.

3. **LibraryManager.cs:**
    *   Manages Visio stencils and shape catalog.
    *   Implements shape operations:
        - Adding shapes with position, size, and color
        - Connecting shapes with straight/curved connectors
        - Text manipulation
        - Shape styling (colors, lines, fills)
        - Grouping and ungrouping
        - Alignment and distribution
        - Property retrieval
    *   Maintains shape categories and stencil organization.
    *   Provides shape catalog to n8n for AI reference.

4. **VisioCommandProcessor.cs:**
    *   Processes JSON commands from n8n.
    *   Supports commands:
        - CreateShape
        - ConnectShapes
        - AddTextToShape
        - SetShapeStyle
        - GroupShapes
        - UngroupShapes
        - AlignShapes
        - DistributeShapes
        - GetShapeProperties
        - GetPageSize
    *   Routes operations to `LibraryManager`.
    *   Handles command validation and error reporting.

5. **ThisAddIn.cs:**
    *   Plugin entry point and initialization.
    *   Manages component lifecycle.
    *   Handles Ribbon UI events.
    *   Sets up webhook listener for n8n communication.
    *   Manages Ollama model availability.

6. **ShapeInfo.cs:**
    *   Manages shape information and properties.
    *   Provides shape details for AI reference.

7. **n8n Workflows:**
    * **Working__Agent_multi_creation.json:**
        - Main workflow for processing chat messages and commands
        - Routes user input through Manager Agent and Action Agent
        - Uses Ollama models for natural language processing
        - Implements JSON schema validation for commands
        - Handles command generation and execution via webhooks
        - Leverages structured output for AI responses.

    * **Image_Agent.json:**
        - Processes image uploads and generates Visio commands
        - Uses Ollama vision model for image analysis
        - Converts image content to structured shape commands
        - Implements auto-fixing output parser for reliable JSON generation
        - Supports both direct image processing and chat-with-image scenarios
        - Outputs structured commands for seamless integration.

    * **Visio_connection_Ollama.json:**
        - Manages connection between Visio and Ollama
        - Lists available Ollama models
        - Provides model information to AIChatPane
        - Handles API communication with Ollama server

    * **Database.json:**
        - Processes Visio stencil catalog
        - Stores shape information in Supabase database
        - Maintains mapping between stencil files and shapes
        - Provides shape catalog data for AI reference
        - Integrated with Supabase for catalog management.