I have a bug for you to help me fix, but first you need to know about my program read the instructions. then read each file. after that and at the end go through the bug i have and try focus on fixing it. 

it should be between the file called StructuredoutputOllama.json and the file called VisioCommandProcessor.cs.

Bug:
Command recieved from n8n, is not being correctly applied in the visio by the visiocommand

Error:
[ProcessWebhookCommand] Received command: {"command":"CreateShape","parameters":{"shapeType":"circle","position":{"x":50,"y":50},"size":{"width":15,"height":15},"color":"red"}}
[ProcessCommand] Received command: {"command":"CreateShape","parameters":{"shapeType":"circle","position":{"x":50,"y":50},"size":{"width":15,"height":15},"color":"red"}}
[ProcessCommand] [Error] 'shapes' array is missing in parameters.
[ProcessWebhookCommand] Command forwarded to VisioCommandProcessor.

Keep in mind that we moved away from tools in the n8n to structured output from ollama.

Example:
Data extraction
To extract structured data from text, define a schema to represent information. The model then extracts the information and returns the data in the defined schema as JSON:

from ollama import chat
from pydantic import BaseModel

class Pet(BaseModel):
  name: str
  animal: str
  age: int
  color: str | None
  favorite_toy: str | None

class PetList(BaseModel):
  pets: list[Pet]

response = chat(
  messages=[
    {
      'role': 'user',
      'content': '''
        I have two pets.
        A cat named Luna who is 5 years old and loves playing with yarn. She has grey fur.
        I also have a 2 year old black cat named Loki who loves tennis balls.
      ''',
    }
  ],
  model='llama3.1',
  format=PetList.model_json_schema(),
)

pets = PetList.model_validate_json(response.message.content)
print(pets)

Example output
pets=[
  Pet(name='Luna', animal='cat', age=5, color='grey', favorite_toy='yarn'), 
  Pet(name='Loki', animal='cat', age=2, color='black', favorite_toy='tennis balls')
]
Image description
Structured outputs can also be used with vision models. For example, the following code uses llama3.2-vision to describe the following image and returns a structured output:

image

from ollama import chat
from pydantic import BaseModel

class Object(BaseModel):
  name: str
  confidence: float
  attributes: str 

class ImageDescription(BaseModel):
  summary: str
  objects: List[Object]
  scene: str
  colors: List[str]
  time_of_day: Literal['Morning', 'Afternoon', 'Evening', 'Night']
  setting: Literal['Indoor', 'Outdoor', 'Unknown']
  text_content: Optional[str] = None

path = 'path/to/image.jpg'

response = chat(
  model='llama3.2-vision',
  format=ImageDescription.model_json_schema(),  # Pass in the schema for the response
  messages=[
    {
      'role': 'user',
      'content': 'Analyze this image and describe what you see, including any objects, the scene, colors and any text you can detect.',
      'images': [path],
    },
  ],
  options={'temperature': 0},  # Set temperature to 0 for more deterministic output
)

image_description = ImageDescription.model_validate_json(response.message.content)
print(image_description)
Example output
summary='A palm tree on a sandy beach with blue water and sky.' 
objects=[
  Object(name='tree', confidence=0.9, attributes='palm tree'), 
  Object(name='beach', confidence=1.0, attributes='sand')
], 
scene='beach', 
colors=['blue', 'green', 'white'], 
time_of_day='Afternoon' 
setting='Outdoor' 
text_content=None

----------


Instructions:
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