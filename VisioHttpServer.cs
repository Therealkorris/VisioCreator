using System;
using System.Threading.Tasks;
using System.Text;
using System.Diagnostics;
using System.Net.Http;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioPlugin
{
    public class VisioCommandSender
    {
        private readonly VisioCommandProcessor commandProcessor;

        public VisioCommandSender(Visio.Application visioApp, LibraryManager libraryManager)
        {
            commandProcessor = new VisioCommandProcessor(visioApp, libraryManager);
        }

        public async Task SendCommandToN8n(string jsonCommand)
        {
            try
            {
                // Debugging: Log that we are entering the method
                Debug.WriteLine($"[Debug] Entering SendCommandToN8n with data: {jsonCommand}");

                string n8nWebhookUrl = "http://localhost:5678/webhook-test/connection_model_list";  // n8n webhook URL

                using (var httpClient = new HttpClient())
                {
                    var content = new StringContent(jsonCommand, Encoding.UTF8, "application/json");
                    // Debugging: Log that we are about to send the POST request
                    Debug.WriteLine($"[Debug] Sending POST request to {n8nWebhookUrl}");

                    var response = await httpClient.PostAsync(n8nWebhookUrl, content);

                    // Debugging: Log if the request was successful
                    if (response.IsSuccessStatusCode)
                    {
                        Debug.WriteLine($"[Debug] Successfully sent data to n8n. Response status: {response.StatusCode}");
                    }
                    else
                    {
                        Debug.WriteLine($"[Debug] Failed to send data to n8n. Response status: {response.StatusCode}, Message: {response.ReasonPhrase}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Debugging: Log if there was an exception
                Debug.WriteLine($"[Error] Exception sending data to n8n: {ex.Message}");
            }
        }


        public async Task ProcessAndSendCommand(string jsonCommand)
        {
            try
            {
                // Process the command locally if needed
                commandProcessor.ProcessCommand(jsonCommand);

                // Send the command to n8n Webhook
                await SendCommandToN8n(jsonCommand);

                Debug.WriteLine("Command processed and sent to n8n.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error processing or sending command: {ex.Message}");
            }
        }
    }
}
