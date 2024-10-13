using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;
using Visio = Microsoft.Office.Interop.Visio;
using System.Diagnostics;

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

        // Function to send Visio command to Python API
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

            // Send the request to the AI Python API
            var response = await httpClient.PostAsync($"{apiEndpoint}/agent-prompt", content);

            // Log the response for debugging
            var responseString = await response.Content.ReadAsStringAsync();
            System.Diagnostics.Debug.WriteLine("Visio Command Response: " + responseString);

            // Process AI response for further action in Visio
            ProcessAICommand(responseString);

            // Return the API response
            return responseString;
        }

        // Function to process the AI response and apply commands to Visio
        public static void ProcessAICommand(string aiResponse)
{
    try
    {
        var commandData = JsonConvert.DeserializeObject<dynamic>(aiResponse);

        // Check if we have a valid action and shape
        if (commandData != null && commandData.action != null)
        {
            string action = commandData.action;
            string shape = commandData.shape;
            float x = (float)commandData.x;
            float y = (float)commandData.y;
            float width = (float)commandData.width;
            float height = (float)commandData.height;
            string color = commandData.color;

            Debug.WriteLine($"Executing action: {action} for shape: {shape}");

            // Depending on the action, call the relevant method in Visio
            if (action == "create_shape")
            {
                // Create an instance of LibraryManager
                var visioApp = Globals.ThisAddIn.Application; // Assuming you're in an add-in
                LibraryManager libraryManager = new LibraryManager(visioApp);
                
                // Use the instance to add the shape to Visio
                libraryManager.AddShapeToDocument("BASIC_M.vssx", shape, x, y, width, height);
            }
            else if (action == "modify_properties")
            {
                // Handle property modifications here (to be implemented)
            }
            // Additional action handling (e.g., connect_shapes) can be added here
        }
        else
        {
            Debug.WriteLine("AI response does not contain a valid action or shape.");
        }
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"Error processing AI command: {ex.Message}");
    }
        }
    }
}
