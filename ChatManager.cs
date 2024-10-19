using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Diagnostics;

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

        public async void SendMessage(string userMessage)
        {
            if (string.IsNullOrEmpty(userMessage)) return;

            appendToChatHistory("User: " + userMessage);

            try
            {
                // Prepare the request content as a multipart form, including the message and the selected model
                var content = new MultipartFormDataContent
        {
            { new StringContent(userMessage), "message" },  // Changed "prompt" to "message"
            { new StringContent(selectedModel), "model" }   // Keeping the selected model
        };

                // Send the message to the n8n webhook for processing
                var response = await httpClient.PostAsync($"{apiEndpoint}/chat-agent", content);
                response.EnsureSuccessStatusCode();  // Ensure the request was successful

                // Read the AI's response
                var responseString = await response.Content.ReadAsStringAsync();

                // Append AI response to chat history
                appendToChatHistory("AI: " + responseString.Trim());

                // Process the AI's response (e.g., if it's a command for Visio)
                await SendCommandToVisio(responseString);
            }
            catch (Exception ex)
            {
                appendToChatHistory("Error: " + ex.Message);
                Debug.WriteLine("Error sending message: " + ex.Message);
            }
        }



        public async void SendMessageWithImage(string imagePath)
        {
            appendToChatHistory("User sent an image: " + System.IO.Path.GetFileName(imagePath));

            try
            {
                var content = new MultipartFormDataContent
                {
                    { new StringContent("Image analysis prompt"), "prompt" },
                    { new StringContent(selectedModel), "model" }
                };

                var imageContent = new ByteArrayContent(System.IO.File.ReadAllBytes(imagePath));
                imageContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("image/jpeg");
                content.Add(imageContent, "file", System.IO.Path.GetFileName(imagePath));

                var response = await httpClient.PostAsync($"{apiEndpoint}/image-prompt", content);

                var responseContent = await response.Content.ReadAsStringAsync();
                appendToChatHistory("AI: " + responseContent);
                Debug.WriteLine($"AI Response: {responseContent}");

                // Process the AI response as a command
                await SendCommandToVisio(responseContent);
            }
            catch (Exception ex)
            {
                appendToChatHistory("Error: " + ex.Message);
                Debug.WriteLine("Error sending image: " + ex.Message);
            }
        }

        private async Task SendCommandToVisio(string aiResponse)
        {
            try
            {
                // Debugging: Log the received AI response
                Debug.WriteLine($"[Debug] Received AI Response: {aiResponse}");

                // Process the command
                var commandProcessor = new VisioCommandProcessor(Globals.ThisAddIn.Application, libraryManager);

                // Debugging: Log before processing the command
                Debug.WriteLine($"[Debug] Processing the command in Visio: {aiResponse}");

                await Task.Run(() => commandProcessor.ProcessCommand(aiResponse));

                // Debugging: Log after processing the command successfully
                Debug.WriteLine($"[Debug] Command processed successfully in Visio.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Error] Error executing Visio command: {ex.Message}");
            }
        }

    }
}
