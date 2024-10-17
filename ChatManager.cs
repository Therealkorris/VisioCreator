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
                var content = new MultipartFormDataContent
                {
                    { new StringContent(userMessage), "prompt" },
                    { new StringContent(selectedModel), "model" }
                };

                var response = await httpClient.PostAsync($"{apiEndpoint}/agent-prompt", content);
                var responseString = await response.Content.ReadAsStringAsync();

                appendToChatHistory("AI: " + responseString.Trim());

                // Send command to the Visio application
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
                // Assuming the AI response is a JSON command
                var commandProcessor = new VisioCommandProcessor(Globals.ThisAddIn.Application, libraryManager);
                await Task.Run(() => commandProcessor.ProcessCommand(aiResponse));
            }
            catch (Exception ex)
            {
                appendToChatHistory("Error executing Visio command: " + ex.Message);
                Debug.WriteLine($"Error executing Visio command: {ex.Message}");
            }
        }
    }
}
