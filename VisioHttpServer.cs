using System;
using System.Net;
using System.Threading.Tasks;
using System.Text;
using System.Diagnostics;
using System.IO;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioPlugin
{
    public class VisioHttpListenerServer
    {
        private readonly HttpListener listener;
        private readonly VisioCommandProcessor commandProcessor;
        private bool isRunning;

        public VisioHttpListenerServer(Visio.Application visioApp, LibraryManager libraryManager)
        {
            listener = new HttpListener();
            listener.Prefixes.Add("http://localhost:8081/"); // Changed port to 8081
            commandProcessor = new VisioCommandProcessor(visioApp, libraryManager);
            isRunning = false;
        }

        public void Start()
        {
            try
            {
                isRunning = true;
                listener.Start();
                Task.Run(() => ListenLoop());
                Debug.WriteLine("Visio HTTP Server started on http://localhost:8081/");
            }
            catch (HttpListenerException ex)
            {
                Debug.WriteLine($"Failed to start HTTP Listener: {ex.Message}");
                isRunning = false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Unexpected error starting HTTP Listener: {ex.Message}");
                isRunning = false;
            }
        }

        public void Stop()
        {
            isRunning = false;
            if (listener.IsListening)
            {
                listener.Stop();
            }
            listener.Close();
            Debug.WriteLine("Visio HTTP Server stopped.");
        }

        private async Task ListenLoop()
        {
            while (isRunning)
            {
                try
                {
                    var context = await listener.GetContextAsync();
                    ProcessRequest(context);
                }
                catch (HttpListenerException ex)
                {
                    if (ex.ErrorCode != 995) // Error code 995 indicates the listener was stopped
                    {
                        Debug.WriteLine($"Listener exception: {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error in ListenLoop: {ex.Message}");
                }
            }
        }

        private void ProcessRequest(HttpListenerContext context)
        {
            var request = context.Request;
            var response = context.Response;

            string responseString = "";
            int statusCode = 200;

            if (request.HttpMethod == "POST" && request.Url.AbsolutePath == "/api/visio/execute")
            {
                try
                {
                    using (var reader = new StreamReader(request.InputStream, request.ContentEncoding))
                    {
                        string jsonCommand = reader.ReadToEnd();
                        if (string.IsNullOrEmpty(jsonCommand))
                        {
                            responseString = "Command cannot be null or empty.";
                            statusCode = 400;
                        }
                        else
                        {
                            commandProcessor.ProcessCommand(jsonCommand);
                            responseString = "{\"status\":\"Success\"}";
                            statusCode = 200;
                        }
                    }
                }
                catch (Exception ex)
                {
                    responseString = $"{{\"status\":\"Error\",\"message\":\"{ex.Message}\"}}";
                    statusCode = 500;
                }
            }
            else
            {
                responseString = "Not Found";
                statusCode = 404;
            }

            byte[] buffer = Encoding.UTF8.GetBytes(responseString);
            response.StatusCode = statusCode;
            response.ContentType = "application/json";
            response.ContentLength64 = buffer.Length;

            try
            {
                response.OutputStream.Write(buffer, 0, buffer.Length);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error writing response: {ex.Message}");
            }
            finally
            {
                response.OutputStream.Close();
            }
        }
    }
}
