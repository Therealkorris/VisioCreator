//using System;
//using System.Threading.Tasks;
//using System.Text;
//using System.Diagnostics;
//using System.Net.Http;
//using Visio = Microsoft.Office.Interop.Visio;
//using Newtonsoft.Json.Linq;

//namespace VisioPlugin
//{
//    public class VisioCommandSender
//    {
//        private readonly VisioCommandProcessor commandProcessor;
//        private readonly string apiEndpoint;

//        // Constructor receives apiEndpoint from ThisAddIn
//        public VisioCommandSender(Visio.Application visioApp, LibraryManager libraryManager, string apiEndpoint)
//        {
//            commandProcessor = new VisioCommandProcessor(visioApp, libraryManager);
//            this.apiEndpoint = apiEndpoint;  // Save the base URL
//        }

//        //public async Task SendCommandToN8n(string jsonCommand)
//        //{
//        //    try
//        //    {
//        //        // Parse the jsonCommand to determine the appropriate n8n endpoint
//        //        JObject commandObject = JObject.Parse(jsonCommand);
//        //        string commandName = commandObject["command"]?.ToString();

//        //        // Debugging: Log that we are entering the method
//        //        Debug.WriteLine($"[Debug] Entering SendCommandToN8n with command: {commandName} and data: {jsonCommand}");

//        //        // Map commands to different n8n endpoints (if necessary)
//        //        string n8nWebhookUrl;
//        //        switch (commandName)
//        //        {
//        //            case "CreateShape":
//        //                n8nWebhookUrl = $"{apiEndpoint}/create_shape";  // Example specific URL for creating shape
//        //                break;

//        //            case "DeleteShape":
//        //                n8nWebhookUrl = $"{apiEndpoint}/delete_shape";  // Example specific URL for deleting shape
//        //                break;

//        //            case "ConnectShapes":
//        //                n8nWebhookUrl = $"{apiEndpoint}/connect_shapes";  // Example specific URL for connecting shapes
//        //                break;

//        //            default:
//        //                n8nWebhookUrl = $"{apiEndpoint}/general_command";  // Default URL for other commands
//        //                break;
//        //        }

//        //        // Debugging: Log that we are about to send the POST request
//        //        Debug.WriteLine($"[Debug] Sending POST request to {n8nWebhookUrl} with command data.");

//        //        using (var httpClient = new HttpClient())
//        //        {
//        //            var content = new StringContent(jsonCommand, Encoding.UTF8, "application/json");

//        //            var response = await httpClient.PostAsync(n8nWebhookUrl, content);

//        //            // Debugging: Log if the request was successful
//        //            if (response.IsSuccessStatusCode)
//        //            {
//        //                Debug.WriteLine($"[Debug] Successfully sent data to n8n. Response status: {response.StatusCode}");
//        //            }
//        //            else
//        //            {
//        //                Debug.WriteLine($"[Debug] Failed to send data to n8n. Response status: {response.StatusCode}, Message: {response.ReasonPhrase}");
//        //            }
//        //        }
//        //    }
//        //    catch (Exception ex)
//        //    {
//        //        // Debugging: Log if there was an exception
//        //        Debug.WriteLine($"[Error] Exception sending data to n8n: {ex.Message}");
//        //    }
//        //}



//        //public async Task ProcessAndSendCommand(string jsonCommand)
//        //{
//        //    try
//        //    {
//        //        // Process the command locally if needed
//        //        commandProcessor.ProcessCommand(jsonCommand);

//        //        // Send the command to n8n Webhook
//        //        await SendCommandToN8n(jsonCommand);

//        //        Debug.WriteLine("Command processed and sent to n8n.");
//        //    }
//        //    catch (Exception ex)
//        //    {
//        //        Debug.WriteLine($"Error processing or sending command: {ex.Message}");
//        //    }
//        //}
//    }
//}
