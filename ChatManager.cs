using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Text;
using System.Net.Http.Headers;
using System.IO;

namespace VisioPlugin
{
    public class VisioChatManager
    {
        public string SelectedModel { get; set; } // Now with a setter!
        private readonly string apiEndpoint;
        private readonly HttpClient httpClient;
        private readonly LibraryManager libraryManager;
        private readonly Action<string> appendToChatHistory;
        private readonly VisioCommandProcessor commandProcessor;
        private readonly AIChatPane chatPane;  // Reference to AIChatPane

        public VisioChatManager(string model, string apiEndpoint, string[] models, LibraryManager libraryManager, Action<string> appendToChatHistory, AIChatPane chatPane)
        {
            this.SelectedModel = model; // Initialize SelectedModel
            this.apiEndpoint = apiEndpoint;
            this.httpClient = new HttpClient();
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
                    model = SelectedModel  // Use the SelectedModel property
                };

                // Convert the payload to JSON
                var jsonContent = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");

                var response = await httpClient.PostAsync($"{apiEndpoint}/chat-agent", jsonContent);
                response.EnsureSuccessStatusCode();

                var responseString = await response.Content.ReadAsStringAsync();

                Debug.WriteLine($"[Debug] Full AI Response (raw): {responseString}");

                await ProcessCommand(responseString, userMessage);  // Await ProcessCommand
            }
            catch (HttpRequestException ex)
            {
                appendToChatHistory("Error sending message (HttpRequestException): " + ex.Message);
                Debug.WriteLine($"[Error] Sending message failed: {ex.Message}");
                chatPane.UpdateCommandStatus(userMessage, "Failed");
            }
            catch (Exception ex)
            {
                appendToChatHistory("Error: " + ex.Message);
                Debug.WriteLine($"[Error] Sending message: {ex.Message}");
                chatPane.UpdateCommandStatus(userMessage, "Failed");
            }
        }

        // Send an image to n8n
        public async Task SendImageToN8n(string imagePath)
        {
            try
            {
                using (var multipartFormContent = new MultipartFormDataContent())
                {
                    // Add the image file
                    var imageStream = File.OpenRead(imagePath);
                    var imageContent = new StreamContent(imageStream);
                    imageContent.Headers.ContentType = MediaTypeHeaderValue.Parse("image/jpeg"); // Or image/png, adjust as needed.
                    multipartFormContent.Add(imageContent, name: "image", fileName: Path.GetFileName(imagePath));

                    // Add the model information as well
                    var modelInfo = new StringContent(SelectedModel, Encoding.UTF8, "text/plain");
                    multipartFormContent.Add(modelInfo, "model");

                    // Send the request to the /chat-agent webhook, which now handles image uploads
                    var response = await httpClient.PostAsync($"{apiEndpoint}/chat-agent", multipartFormContent);
                    response.EnsureSuccessStatusCode();

                    var responseString = await response.Content.ReadAsStringAsync();
                    Debug.WriteLine($"[SendImageToN8n] Response: {responseString}");

                    // Process the response as before (chat message or command)
                    await ProcessCommand(responseString, $"Image: {Path.GetFileName(imagePath)}");

                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SendImageToN8n] Error sending image: {ex.Message}");
                appendToChatHistory($"Error sending image: {ex.Message}");
            }
        }

        // Process the AI's response and decide if it's a chat message or a command
        private async Task ProcessCommand(string aiResponse, string userMessage)
        {
            try
            {
                // Log the raw response again before processing
                Debug.WriteLine($"[Debug] Received AI Response (raw): {aiResponse}");

                // Check if the response is empty or invalid
                if (string.IsNullOrEmpty(aiResponse))
                {
                    Debug.WriteLine("[Error] Received empty AI response.");
                    appendToChatHistory("[Error] Received empty response from AI.");

                    // Update command status to Failed
                    chatPane.UpdateCommandStatus(userMessage, "Failed");
                    return;
                }

                // Validate that the AI response is in JSON format
                if (IsValidJson(aiResponse))
                {
                    JObject responseObject = JObject.Parse(aiResponse);
                    Debug.WriteLine($"[Debug] Parsed JSON Response: {responseObject}");

                    // If the response contains a command, execute it in Visio
                    if (responseObject["command"] != null)
                    {
                        Debug.WriteLine($"[Debug] Command found: {responseObject["command"]}");

                        // Send the parsed command to VisioCommandProcessor to execute the action in Visio
                        await Task.Run(() => commandProcessor.ProcessCommand(aiResponse));

                        Debug.WriteLine($"[Debug] Command executed in Visio.");

                        // Update command status to Success
                        chatPane.UpdateCommandStatus(userMessage, "Success");
                    }
                    else if (responseObject["message"] != null)
                    {
                        // If the response is a regular chat message, append it to the chat history
                        string chatMessage = responseObject["message"].ToString();
                        appendToChatHistory($"AI: {chatMessage}");
                    }
                    else
                    {
                        Debug.WriteLine("[Error] Unrecognized response format.");
                        chatPane.UpdateCommandStatus(userMessage, "Failed");
                    }
                }
                else
                {
                    // If the response is not valid JSON, treat it as plain text chat message
                    Debug.WriteLine("[Error] AI Response is not valid JSON. Treating it as plain text.");
                    appendToChatHistory($"AI: {aiResponse}");

                    // Update command status to Success (since plain text is not considered a failure)
                    chatPane.UpdateCommandStatus(userMessage, "Success");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Error] Error processing AI response: {ex.Message}");

                // Update command status to Failed
                chatPane.UpdateCommandStatus(userMessage, "Failed");
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
