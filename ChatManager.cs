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

        public VisioChatManager(string model, string apiEndpoint, string[] models, LibraryManager libraryManager, Action<string> appendToChatHistory)
        {
            this.selectedModel = model;
            this.apiEndpoint = apiEndpoint;
            this.httpClient = new HttpClient();
            this.libraryManager = libraryManager;
            this.appendToChatHistory = appendToChatHistory;
            this.commandProcessor = new VisioCommandProcessor(Globals.ThisAddIn.Application, libraryManager);
        }

        // Send a message to the API and process the response
        public async void SendMessage(string userMessage)
        {
            if (string.IsNullOrEmpty(userMessage)) return;

            appendToChatHistory("User: " + userMessage); // Append the user's message

            try
            {
                var payload = new
                {
                    message = userMessage,
                    model = selectedModel  // Model specified for AI processing
                };

                // Convert the payload to JSON
                var jsonContent = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");

                // Send the message to the external API (e.g., n8n or another service)
                var response = await httpClient.PostAsync($"{apiEndpoint}/chat-agent", jsonContent);
                response.EnsureSuccessStatusCode();  // Ensure the API call succeeded

                // Read the response string
                var responseString = await response.Content.ReadAsStringAsync();

                // Process the AI response (e.g., if it's a command for Visio)
                await ProcessCommandResponse(responseString);
            }
            catch (HttpRequestException ex)
            {
                appendToChatHistory("Error sending message (HttpRequestException): " + ex.Message);
                Debug.WriteLine($"[Error] Sending message failed: {ex.Message}");
            }
            catch (Exception ex)
            {
                appendToChatHistory("Error: " + ex.Message);
                Debug.WriteLine($"[Error] Sending message: {ex.Message}");
            }
        }

        // Process the received AI response and send it to Visio
        private async Task ProcessCommandResponse(string aiResponse)
        {
            try
            {
                Debug.WriteLine($"[Debug] Received AI Response: {aiResponse}");

                // Validate that the AI response is in JSON format
                if (IsValidJson(aiResponse))
                {
                    JObject responseObject = JObject.Parse(aiResponse);
                    Debug.WriteLine($"[Debug] Parsed JSON Response: {responseObject}");

                    // If the response contains a command, process it
                    if (responseObject["command"] != null)
                    {
                        Debug.WriteLine($"[Debug] Command found: {responseObject["command"]}");
                        await SendCommandToVisio(aiResponse); // Sends the command to Visio
                    }
                    else if (responseObject["message"] != null)
                    {
                        // If the response is a regular message, append it to the chat history
                        string chatMessage = responseObject["message"].ToString();
                        appendToChatHistory($"AI: {chatMessage}");
                    }
                    else
                    {
                        Debug.WriteLine("[Error] Unrecognized response format.");
                    }
                }
                else
                {
                    Debug.WriteLine("[Error] AI Response is not valid JSON. Treating it as plain text.");
                    appendToChatHistory($"AI: {aiResponse}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Error] Error processing AI response: {ex.Message}");
            }
        }

        // Method to send command to Visio
        public async Task SendCommandToVisio(string jsonCommand)
        {
            try
            {
                Debug.WriteLine($"[Debug] Sending command to Visio: {jsonCommand}");

                // Prepare the command in JSON format
                var jsonContent = new StringContent(jsonCommand, Encoding.UTF8, "application/json");

                // Send the request to the Visio command API
                var response = await httpClient.PostAsync($"{apiEndpoint}/visio-command", jsonContent);

                Debug.WriteLine($"[Debug] Visio API Response Status: {response.StatusCode}");

                if (!response.IsSuccessStatusCode)
                {
                    Debug.WriteLine($"[Error] Visio API returned error: {response.StatusCode}");
                    appendToChatHistory($"Error: Visio API returned status {response.StatusCode}");
                    return;
                }

                // Read and log the response content
                var responseString = await response.Content.ReadAsStringAsync();
                Debug.WriteLine($"[Debug] Visio API Response Content: {responseString}");
            }
            catch (HttpRequestException ex)
            {
                Debug.WriteLine($"[Error] HTTP Request failed: {ex.Message}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Error] General error: {ex.Message}");
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
