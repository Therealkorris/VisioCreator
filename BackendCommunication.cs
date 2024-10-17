using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Diagnostics;
using Newtonsoft.Json;
using System.Linq;
using System.Text;

namespace VisioPlugin
{
    public static class BackendCommunication
    {
        private static readonly HttpClient httpClient = new HttpClient();

        /// <summary>
        /// Gets the list of available models from the backend API.
        /// </summary>
        /// <param name="apiEndpoint">The base URL of the API endpoint.</param>
        /// <returns>An array of model names.</returns>
        public static async Task<string[]> GetModels(string apiEndpoint)
        {
            try
            {
                var response = await httpClient.GetAsync($"{apiEndpoint}/models");
                response.EnsureSuccessStatusCode();
                var content = await response.Content.ReadAsStringAsync();

                var jsonResponse = JsonConvert.DeserializeObject<ModelResponse>(content);
                if (jsonResponse?.Models == null || !jsonResponse.Models.Any())
                    throw new Exception("No models found in the API response");

                return jsonResponse.Models.Select(m => m.Name).ToArray();
            }
            catch (HttpRequestException e)
            {
                Debug.WriteLine($"HTTP Request Error: {e.Message}");
                throw;
            }
            catch (Exception e)
            {
                Debug.WriteLine($"Unexpected error: {e.Message}");
                throw;
            }
        }

        /// <summary>
        /// Sends a chat message to the backend AI system and receives the response.
        /// </summary>
        /// <param name="apiEndpoint">The base URL of the API endpoint.</param>
        /// <param name="message">The user's message to send.</param>
        /// <param name="model">The AI model to use.</param>
        /// <returns>The AI's response message.</returns>
        public static async Task<string> SendMessage(string apiEndpoint, string message, string model)
        {
            try
            {
                var payload = new
                {
                    model = model,
                    message = message
                };
                var jsonContent = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
                var response = await httpClient.PostAsync($"{apiEndpoint}/agent-prompt", jsonContent);
                response.EnsureSuccessStatusCode();
                var responseContent = await response.Content.ReadAsStringAsync();

                var chatResponse = JsonConvert.DeserializeObject<ChatResponse>(responseContent);
                if (chatResponse == null || string.IsNullOrEmpty(chatResponse.Response))
                    throw new Exception("Invalid response from AI system");

                return chatResponse.Response;
            }
            catch (HttpRequestException e)
            {
                Debug.WriteLine($"HTTP Request Error: {e.Message}");
                throw;
            }
            catch (Exception e)
            {
                Debug.WriteLine($"Unexpected error: {e.Message}");
                throw;
            }
        }

        /// <summary>
        /// Sends an image along with a message to the backend AI system.
        /// </summary>
        /// <param name="apiEndpoint">The base URL of the API endpoint.</param>
        /// <param name="imagePath">The path to the image file.</param>
        /// <param name="model">The AI model to use.</param>
        /// <returns>The AI's response message.</returns>
        public static async Task<string> SendImage(string apiEndpoint, string imagePath, string model)
        {
            try
            {
                var content = new MultipartFormDataContent();

                var imageContent = new ByteArrayContent(System.IO.File.ReadAllBytes(imagePath));
                imageContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("image/jpeg");
                content.Add(imageContent, "image", System.IO.Path.GetFileName(imagePath));

                content.Add(new StringContent(model), "model");

                var response = await httpClient.PostAsync($"{apiEndpoint}/image", content);
                response.EnsureSuccessStatusCode();
                var responseContent = await response.Content.ReadAsStringAsync();

                var imageResponse = JsonConvert.DeserializeObject<ChatResponse>(responseContent);
                if (imageResponse == null || string.IsNullOrEmpty(imageResponse.Response))
                    throw new Exception("Invalid response from AI system");

                return imageResponse.Response;
            }
            catch (HttpRequestException e)
            {
                Debug.WriteLine($"HTTP Request Error: {e.Message}");
                throw;
            }
            catch (Exception e)
            {
                Debug.WriteLine($"Unexpected error: {e.Message}");
                throw;
            }
        }

        // Models to deserialize API responses
        public class ModelResponse
        {
            public System.Collections.Generic.List<ModelInfo> Models { get; set; }
        }

        public class ModelInfo
        {
            public string Name { get; set; }
        }

        public class ChatResponse
        {
            public string Response { get; set; }
        }
    }
}
