using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Text;
using System.Net.Http.Headers;
using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace VisioPlugin
{
    public class VisioChatManager
    {
        public string SelectedModel { get; set; } // Now with a setter!
        private readonly string apiEndpoint;
        private static readonly HttpClient httpClient = new HttpClient() { Timeout = TimeSpan.FromMinutes(30) };
        private readonly LibraryManager libraryManager;
        private readonly Action<string> appendToChatHistory;
        private readonly VisioCommandProcessor commandProcessor;
        private readonly AIChatPane chatPane;  // Reference to AIChatPane

        public VisioChatManager(string model, string apiEndpoint, string[] models, LibraryManager libraryManager, Action<string> appendToChatHistory, AIChatPane chatPane)
        {
            this.SelectedModel = model; // Initialize SelectedModel
            this.apiEndpoint = apiEndpoint;
            this.libraryManager = libraryManager;
            this.appendToChatHistory = appendToChatHistory;
            this.commandProcessor = new VisioCommandProcessor(Globals.ThisAddIn.Application, libraryManager);
            this.chatPane = chatPane;
        }

        // Send a message to the AI and process the response (chat or command)
        public async Task SendMessage(string userMessage)  // Return Task
        {
            if (string.IsNullOrEmpty(userMessage)) return;
            try
            {
                var payload = new
                {
                    message = userMessage,
                    model = SelectedModel
                };

                var jsonContent = new StringContent(
                    Newtonsoft.Json.JsonConvert.SerializeObject(payload), 
                    Encoding.UTF8, 
                    "application/json"
                );

                var response = await httpClient.PostAsync($"{apiEndpoint}/chat-agent", jsonContent);
                response.EnsureSuccessStatusCode();

                var responseString = await response.Content.ReadAsStringAsync();
                Debug.WriteLine($"[Debug] Full AI Response (raw): {responseString}");

                // Process the command and wait for it to complete
                await ProcessWebhookCommand(responseString, userMessage);
            }
            catch (HttpRequestException ex)
            {
                appendToChatHistory("Error sending message (HttpRequestException): " + ex.Message);
                Debug.WriteLine($"[Error] Sending message failed: {ex.Message}");
                chatPane.UpdateCommandStatus(new CommandDetails
                {
                    Command = "Error",
                    Status = "Failed",
                    UserMessage = userMessage,
                    AIResponse = ex.Message
                });
            }
            catch (Exception ex)
            {
                appendToChatHistory("Error: " + ex.Message);
                Debug.WriteLine($"[Error] Sending message: {ex.Message}");
                var commandDetails = chatPane.GetCurrentCommand();
                if (commandDetails != null)
                {
                    commandDetails.Status = "Failed";
                    commandDetails.AIResponse = ex.Message;
                    chatPane.UpdateCommandStatus(commandDetails);
                    chatPane.ResetCurrentCommand();
                }
            }
        }

        private async Task ProcessWebhookCommand(string aiResponse, string userMessage)
        {
            Debug.WriteLine("[ProcessWebhookCommand] Command forwarded to VisioCommandProcessor.");
            
            try
            {
                // Get the current command before processing
                var commandDetails = chatPane.GetCurrentCommand();
                if (commandDetails == null)
                {
                    Debug.WriteLine("[ProcessWebhookCommand] No active command found.");
                    return;
                }

                // First, try to process it as a Visio command if it's valid JSON
                if (IsValidJson(aiResponse))
                {
                    var commandObject = JObject.Parse(aiResponse);
                    if (commandObject["command"] != null)
                    {
                        // Execute the command and wait for it to complete
                        await Task.Run(() => commandProcessor.ProcessCommand(aiResponse));

                        // Get the affected shapes after command execution
                        var affectedShapes = commandProcessor.GetLastAffectedShapes().ToList();
                        Debug.WriteLine($"[ProcessWebhookCommand] Retrieved {affectedShapes.Count} affected shapes");

                        // Update affected shapes in command details
                        commandDetails.AffectedShapeIds = affectedShapes.Select(s => s.ID16.ToString()).ToList();
                    }
                }

                // Extract chat message from AI response
                string chatMessage;
                if (IsValidJson(aiResponse))
                {
                    var responseObject = JObject.Parse(aiResponse);
                    chatMessage = ExtractChatMessage(responseObject);
                }
                else
                {
                    chatMessage = aiResponse;
                }

                // Update command details
                commandDetails.Status = "Success";
                commandDetails.AIResponse = chatMessage;
                commandDetails.VisioCommand = aiResponse;
                chatPane.UpdateCommandStatus(commandDetails);

                // Add the AI response to chat history
                appendToChatHistory($"AI: {chatMessage}");

                // Reset the current command after successful processing
                chatPane.ResetCurrentCommand();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessWebhookCommand] Error: {ex.Message}");
                var commandDetails = chatPane.GetCurrentCommand();
                if (commandDetails != null)
                {
                    commandDetails.Status = "Failed";
                    commandDetails.AIResponse = $"Error: {ex.Message}";
                    chatPane.UpdateCommandStatus(commandDetails);
                    chatPane.ResetCurrentCommand();
                }
                throw;
            }
        }

        // Send an image to n8n
        public async Task SendImageToN8n(string imagePath, string userMessage = null)
        {
            try
            {
                using (var multipartFormContent = new MultipartFormDataContent())
                {
                    // Open the image file as a stream
                    var imageStream = File.OpenRead(imagePath);

                    // Create StreamContent for the image and set the Content-Type
                    var imageContent = new StreamContent(imageStream);
                    imageContent.Headers.ContentType = MediaTypeHeaderValue.Parse("image/jpeg"); // Adjust MIME type if needed

                    // Add the image content to the form
                    multipartFormContent.Add(imageContent, "image", Path.GetFileName(imagePath));

                    // Add the selected model as a separate form field
                    var modelInfo = new StringContent(SelectedModel, Encoding.UTF8, "text/plain");
                    multipartFormContent.Add(modelInfo, "model");

                    // Add the user message if provided
                    if (!string.IsNullOrEmpty(userMessage))
                    {
                        var messageContent = new StringContent(userMessage, Encoding.UTF8, "text/plain");
                        multipartFormContent.Add(messageContent, "message");
                    }

                    Debug.WriteLine($"[SendImageToN8n] Sending image: {Path.GetFileName(imagePath)} with model {SelectedModel} and message: {userMessage} to {apiEndpoint}/image-agent");

                    // Send the POST request to the /image-agent endpoint
                    var response = await httpClient.PostAsync($"{apiEndpoint}/image-agent", multipartFormContent);
                    response.EnsureSuccessStatusCode();

                    // Read the response content
                    var responseString = await response.Content.ReadAsStringAsync();
                    Debug.WriteLine($"[SendImageToN8n] Response: {responseString}");

                    // Process the AI response
                    await ProcessCommand(responseString, $"Image: {Path.GetFileName(imagePath)}" + (!string.IsNullOrEmpty(userMessage) ? $" - {userMessage}" : ""));
                }
            }
            catch (HttpRequestException ex)
            {
                Debug.WriteLine($"[SendImageToN8n] HttpRequestException: {ex.Message}");
                appendToChatHistory($"Error sending image: {ex.Message}");
                chatPane.UpdateCommandStatus(new CommandDetails
                {
                    Command = $"Image: {Path.GetFileName(imagePath)}",
                    Status = "Failed",
                    UserMessage = userMessage,
                    AIResponse = ex.Message
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SendImageToN8n] Exception: {ex.Message}");
                appendToChatHistory($"Error sending image: {ex.Message}");
                chatPane.UpdateCommandStatus(new CommandDetails
                {
                    Command = $"Image: {Path.GetFileName(imagePath)}",
                    Status = "Failed",
                    UserMessage = userMessage,
                    AIResponse = ex.Message
                });
            }
        }


        // Process the AI's response and decide if it's a chat message or a command
        private async Task ProcessCommand(string aiResponse, string userMessage)
        {
            try
            {
                Debug.WriteLine("\n[ProcessCommand] ========== START COMMAND PROCESSING ==========");
                Debug.WriteLine($"[ProcessCommand] User Message: {userMessage}");
                Debug.WriteLine($"[ProcessCommand] Raw AI Response: {aiResponse}");

                if (string.IsNullOrEmpty(aiResponse))
                {
                    Debug.WriteLine("[ProcessCommand] ERROR: Received empty AI response.");
                    appendToChatHistory("[Error] Received empty response from AI.");
                    chatPane.UpdateCommandStatus(new CommandDetails
                    {
                        Command = "Error",
                        Status = "Failed",
                        UserMessage = userMessage,
                        AIResponse = "Empty response"
                    });
                    return;
                }

                // Get the existing command
                var commandDetails = chatPane.GetCurrentCommand();
                if (commandDetails == null)
                {
                    Debug.WriteLine("[ProcessCommand] No existing command found, this shouldn't happen!");
                    return;
                }

                Debug.WriteLine("[ProcessCommand] Getting affected shapes...");
                var affectedShapes = await Task.Run(() => commandProcessor.GetLastAffectedShapes().ToList());
                Debug.WriteLine($"[ProcessCommand] Retrieved {affectedShapes.Count} affected shapes");

                var affectedShapeIds = affectedShapes
                    .Where(shape => shape != null)
                    .Select(shape => {
                        Debug.WriteLine($"[ProcessCommand] Processing shape with ID: {shape.ID16}");
                        return shape.ID16.ToString();
                    })
                    .Distinct()
                    .ToList();

                Debug.WriteLine($"[ProcessCommand] Final affected shape IDs: {string.Join(", ", affectedShapeIds)}");

                // Extract the AI response
                string chatMessage = "";
                if (IsValidJson(aiResponse))
                {
                    Debug.WriteLine("[ProcessCommand] AI Response is valid JSON, parsing...");
                    JObject responseObject = JObject.Parse(aiResponse);
                    chatMessage = ExtractChatMessage(responseObject);
                    Debug.WriteLine($"[ProcessCommand] Extracted chat message: {chatMessage}");
                    appendToChatHistory($"AI: {chatMessage}");
                }
                else
                {
                    Debug.WriteLine("[ProcessCommand] AI Response is not valid JSON, using as is");
                    chatMessage = aiResponse;
                    appendToChatHistory($"AI: {aiResponse}");
                }

                // Update the existing command with all details
                Debug.WriteLine("[ProcessCommand] Updating command details...");
                commandDetails.Status = "Success";
                commandDetails.AIResponse = chatMessage;
                commandDetails.VisioCommand = aiResponse;
                commandDetails.AffectedShapeIds = affectedShapeIds;

                Debug.WriteLine("\n[ProcessCommand] ========== FINAL COMMAND DETAILS ==========");
                Debug.WriteLine($"[ProcessCommand] Command: {commandDetails.Command}");
                Debug.WriteLine($"[ProcessCommand] Status: {commandDetails.Status}");
                Debug.WriteLine($"[ProcessCommand] User Message: {commandDetails.UserMessage}");
                Debug.WriteLine($"[ProcessCommand] AI Response: {commandDetails.AIResponse}");
                Debug.WriteLine($"[ProcessCommand] Affected Shapes: {string.Join(", ", commandDetails.AffectedShapeIds)}");
                Debug.WriteLine($"[ProcessCommand] Visio Command: {commandDetails.VisioCommand}");
                
                Debug.WriteLine("[ProcessCommand] Updating command status...");
                chatPane.UpdateCommandStatus(commandDetails);
                Debug.WriteLine("[ProcessCommand] Command status updated");
                Debug.WriteLine("[ProcessCommand] ========== END COMMAND PROCESSING ==========\n");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessCommand] ========== ERROR ==========");
                Debug.WriteLine($"[ProcessCommand] Error Type: {ex.GetType().Name}");
                Debug.WriteLine($"[ProcessCommand] Error Message: {ex.Message}");
                Debug.WriteLine($"[ProcessCommand] Stack Trace: {ex.StackTrace}");
                
                var commandDetails = chatPane.GetCurrentCommand();
                if (commandDetails != null)
                {
                    commandDetails.Status = "Failed";
                    commandDetails.AIResponse = ex.Message;
                    chatPane.UpdateCommandStatus(commandDetails);
                }
                else
                {
                    chatPane.UpdateCommandStatus(new CommandDetails
                    {
                        Command = "Error",
                        Status = "Failed",
                        UserMessage = userMessage,
                        AIResponse = ex.Message
                    });
                }
                Debug.WriteLine("[ProcessCommand] ========== END ERROR ==========\n");
            }
        }

        private string ExtractChatMessage(JObject responseObject)
        {
            try
            {
                // Get the first property which contains the main message
                var firstProperty = responseObject.Properties().FirstOrDefault();
                if (firstProperty == null) return "";

                // If the value is a simple string, return it
                if (firstProperty.Value.Type == JTokenType.String)
                    return firstProperty.Value.ToString();

                // If it's a nested structure, try to find the deepest string value
                var message = firstProperty.Name;
                var currentToken = firstProperty.Value;

                // Try to find the actual message in the nested structure
                while (currentToken is JObject obj)
                {
                    var emptyKeyProp = obj.Properties().FirstOrDefault(p => p.Name == "");
                    if (emptyKeyProp != null && emptyKeyProp.Value.Type == JTokenType.String)
                    {
                        return emptyKeyProp.Value.ToString();
                    }

                    var firstProp = obj.Properties().FirstOrDefault();
                    if (firstProp == null) break;

                    if (!string.IsNullOrWhiteSpace(firstProp.Name) && firstProp.Name != " ")
                    {
                        message = firstProp.Name;
                    }

                    currentToken = firstProp.Value;
                }

                // Clean up the message
                message = message.Replace("\n", " ").Trim();
                Debug.WriteLine($"[ExtractChatMessage] Extracted message: {message}");
                return message;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExtractChatMessage] Error extracting message: {ex.Message}");
                return "";
            }
        }

        // Validate if the input string is a valid JSON object
        private bool IsValidJson(string strInput)
        {
            strInput = strInput.Trim();
            if ((strInput.StartsWith("{") && strInput.EndsWith("}")) ||  // Object check
                (strInput.StartsWith("[") && strInput.EndsWith("]")))   // Array check
            {
                try
                {
                    var obj = JToken.Parse(strInput);  // Try parsing the string into a JSON object
                    return true;
                }
                catch (Exception)
                {
                    return false;  // If parsing fails, return false
                }
            }
            return false;
        }
    }
}