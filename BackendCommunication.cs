using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;

namespace VisioPlugin
{
    public static class BackendCommunication
    {
        private static readonly HttpClient httpClient = new HttpClient();

        // Function to send text message to AI API
        public static async Task<string> SendTextMessage(string message, string model, string apiEndpoint)
        {
            var content = new MultipartFormDataContent();
            content.Add(new StringContent(message), "prompt");
            content.Add(new StringContent(model), "model");

            var response = await httpClient.PostAsync($"{apiEndpoint}/text-prompt", content);
            return await response.Content.ReadAsStringAsync();
        }

        // Function to send image message to AI API
        public static async Task<string> SendImageMessage(string imagePath, string model, string apiEndpoint)
        {
            var content = new MultipartFormDataContent();
            content.Add(new StringContent(model), "model");

            var imageContent = new ByteArrayContent(File.ReadAllBytes(imagePath));
            imageContent.Headers.ContentType = new MediaTypeHeaderValue("image/jpeg");
            content.Add(imageContent, "file", Path.GetFileName(imagePath));

            var response = await httpClient.PostAsync($"{apiEndpoint}/image-prompt", content);
            return await response.Content.ReadAsStringAsync();
        }

        // Function to send Visio command to Python API (New function)
        public static async Task<string> SendVisioCommand(string action, string shape, float x, float y, float width, float height, string apiEndpoint)
        {
            var command = new
            {
                action = action,
                shape = shape,
                x = x,
                y = y,
                width = width,
                height = height
            };

            // Serialize the command into JSON format
            string jsonCommand = JsonConvert.SerializeObject(command);
            var content = new StringContent(jsonCommand, Encoding.UTF8, "application/json");

            // Send the request to the Python API
            var response = await httpClient.PostAsync($"{apiEndpoint}/test-visio-command", content);

            // Log the response for debugging
            var responseString = await response.Content.ReadAsStringAsync();
            System.Diagnostics.Debug.WriteLine("Visio Command Response: " + responseString);

            // Return the API response
            return responseString;
        }
    }
}
