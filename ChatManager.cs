using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Text;

namespace VisioPlugin
{
    public class VisioChatManager
    {
        private string selectedModel;
        private readonly string apiEndpoint;
        private readonly HttpClient httpClient;
        private readonly LibraryManager libraryManager;
        private readonly Action<string> appendToChatHistory;
        private readonly VisioCommandProcessor commandProcessor;
        private readonly AIChatPane chatPane;  // Reference to AIChatPane

        public VisioChatManager(string model, string apiEndpoint, string[] models, LibraryManager libraryManager, Action<string> appendToChatHistory, AIChatPane chatPane)
        {
            this.selectedModel = model;
            this.apiEndpoint = apiEndpoint;
            this.httpClient = new HttpClient();
            this.libraryManager = libraryManager;
            this.appendToChatHistory = appendToChatHistory;
            this.commandProcessor = new VisioCommandProcessor(Globals.ThisAddIn.Application, libraryManager);
            this.chatPane = chatPane;  // Store the reference to AIChatPane
        }

        // Send a message to the AI and process the response (chat or command)
        public async void SendMessage(string userMessage)
        {
            Debug.WriteLine($"[SendMessage] Sending message to AI: {userMessage}");
            if (string.IsNullOrEmpty(userMessage)) return;
            try
            {
                var payload = new
                {
                    message = userMessage,
                    model = selectedModel  // Model specified for AI processing
                };

                // Convert the payload to JSON
                var jsonContent = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");

                // Send the message to the external API
                Debug.WriteLine($"[SendMessage] Sending request to API endpoint: {apiEndpoint}/chat-agent");

                var response = await httpClient.PostAsync($"{apiEndpoint}/chat-agent", jsonContent);

                Debug.WriteLine($"[SendMessage] API response status: {response.StatusCode}");

                response.EnsureSuccessStatusCode();  // Ensure the API call succeeded

                // Read the response from AI
                var responseString = await response.Content.ReadAsStringAsync();

                // Log the full raw response for debugging purposes
                Debug.WriteLine($"[SendMessage] Full AI Response (raw): {responseString}");

                // Process AI response - either chat or a command
                await ProcessCommand(responseString, userMessage);
            }
            catch (HttpRequestException ex)
            {
                appendToChatHistory("Error sending message (HttpRequestException): " + ex.Message);
                Debug.WriteLine($"[SendMessage] [Error] Sending message failed: {ex.Message}");

                // Update command status to Failed
                chatPane.UpdateCommandStatus(userMessage, "Failed");
            }
            catch (Exception ex)
            {
                appendToChatHistory("Error: " + ex.Message);
                Debug.WriteLine($"[SendMessage] [Error] Sending message: {ex.Message}");

                // Update command status to Failed
                chatPane.UpdateCommandStatus(userMessage, "Failed");
            }
        }

        // Process the AI's response and decide if it's a chat message or a command
        private async Task ProcessCommand(string aiResponse, string userMessage)
        {
            Debug.WriteLine($"[ProcessCommand] Processing AI response: {aiResponse}");
            try
            {
                // Log the raw response again before processing
                Debug.WriteLine($"[ProcessCommand] Received AI Response (raw): {aiResponse}");

                // Check if the response is empty or invalid
                if (string.IsNullOrEmpty(aiResponse))
                {
                    Debug.WriteLine("[ProcessCommand] [Error] Received empty AI response.");
                    appendToChatHistory("[ProcessCommand] [Error] Received empty response from AI.");

                    // Update command status to Failed
                    chatPane.UpdateCommandStatus(userMessage, "Failed");
                    return;
                }

                // Validate that the AI response is in JSON format
                if (IsValidJson(aiResponse))
                {
                    JObject responseObject = JObject.Parse(aiResponse);
                    Debug.WriteLine($"[ProcessCommand] Parsed JSON Response: {responseObject}");

                    // If the response contains a command, execute it in Visio
                    if (responseObject["command"] != null)
                    {
                        Debug.WriteLine($"[ProcessCommand] Command found: {responseObject["command"]}");

                        // Send the parsed command to VisioCommandProcessor to execute the action in Visio
                        await Task.Run(() => commandProcessor.ProcessCommand(aiResponse));

                        Debug.WriteLine($"[ProcessCommand] Command executed in Visio.");

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
                        Debug.WriteLine("[ProcessCommand] [Error] Unrecognized response format.");
                        chatPane.UpdateCommandStatus(userMessage, "Failed");
                    }
                }
                else
                {
                    // If the response is not valid JSON, treat it as plain text chat message
                    Debug.WriteLine("[ProcessCommand] [Error] AI Response is not valid JSON. Treating it as plain text.");
                    appendToChatHistory($"AI: {aiResponse}");

                    // Update command status to Success (since plain text is not considered a failure)
                    chatPane.UpdateCommandStatus(userMessage, "Success");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessCommand] [Error] Error processing AI response: {ex.Message}");

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
