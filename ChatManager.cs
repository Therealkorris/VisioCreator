﻿using System;
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

        // Send a message to the AI and process the response (chat or command)
        public async void SendMessage(string userMessage)
        {
            if (string.IsNullOrEmpty(userMessage)) return;

            appendToChatHistory("User: " + userMessage); // Append the user's message to chat

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
                var response = await httpClient.PostAsync($"{apiEndpoint}/chat-agent", jsonContent);
                response.EnsureSuccessStatusCode();  // Ensure the API call succeeded

                // Read the response from AI
                var responseString = await response.Content.ReadAsStringAsync();

                // Process AI response - either chat or a command
                await ProcessCommand(responseString);
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

        // Process the AI's response and decide if it's a chat message or a command
        private async Task ProcessCommand(string aiResponse)
        {
            try
            {
                Debug.WriteLine($"[Debug] Received AI Response: {aiResponse}");

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
                    }
                }
                else
                {
                    // If the response is not a valid JSON, treat it as plain text chat message
                    Debug.WriteLine("[Error] AI Response is not valid JSON. Treating it as plain text.");
                    appendToChatHistory($"AI: {aiResponse}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Error] Error processing AI response: {ex.Message}");
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
