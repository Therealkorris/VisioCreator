using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;
using System.Windows.Forms;
using System.Net.Http;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Collections.Concurrent;

namespace VisioPlugin
{
    [ComVisible(true)]
    public partial class ThisAddIn
    {
        private Visio.Application visioApplication;
        internal Office.IRibbonUI Ribbon { get; set; }
        private LibraryManager libraryManager;
        private System.Windows.Forms.Control uiControl;
        internal string CurrentCategory { get; set; }

        private string _apiEndpoint = "http://localhost:5678/webhook";
        public string apiEndpoint
        {
            get => _apiEndpoint;
            set
            {
                try
                {
                    // Ensure proper URL format
                    string formattedUrl = value;
                    if (!formattedUrl.StartsWith("http://") && !formattedUrl.StartsWith("https://"))
                    {
                        formattedUrl = "http://" + formattedUrl;
                    }

                    if (Uri.TryCreate(formattedUrl, UriKind.Absolute, out Uri uri))
                    {
                        _apiEndpoint = formattedUrl;
                        ApiConfig.UpdateFromApiEndpoint(formattedUrl);
                        Debug.WriteLine($"[ThisAddIn] API Endpoint updated to: {formattedUrl}");
                    }
                    else
                    {
                        Debug.WriteLine($"[ThisAddIn] Invalid API Endpoint format: {value}");
                        MessageBox.Show("Please enter a valid URL (e.g., http://localhost:5678/webhook)");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ThisAddIn] Error updating API Endpoint: {ex.Message}");
                    MessageBox.Show($"Error updating API Endpoint: {ex.Message}");
                }
            }
        }

        public string GetAPIEndpointText(Office.IRibbonControl control)
        {
            return _apiEndpoint;
        }

        public bool isConnected = false;
        private string[] availableModels = new string[0];
        private HttpClient httpClient = new HttpClient();
        private string selectedModel = "phi-4";
        private AIChatPane aiChatPane;
        private VisioCommandProcessor commandProcessor;
        private HttpListener listener;
        private VisioChatManager visioChatManager;
        private ConcurrentDictionary<string, HttpListenerContext> activeConnections = new ConcurrentDictionary<string, HttpListenerContext>();
        private const int VISIO_PORT = 5680;
        private const int N8N_PORT = 5678;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonExtension(this);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Set default API configuration
                ApiConfig.UpdateFromApiEndpoint("http://localhost:5678/webhook");
                
                Debug.WriteLine("Initializing Visio application...");
                visioApplication = (Visio.Application)Application;

                Debug.WriteLine("Initializing LibraryManager...");
                libraryManager = new LibraryManager(visioApplication);

                Debug.WriteLine("Initializing UIControl...");
                uiControl = new System.Windows.Forms.Control();
                uiControl.CreateControl();

                commandProcessor = new VisioCommandProcessor(visioApplication, libraryManager);

                Debug.WriteLine("Initializing VisioChatManager...");
                visioChatManager = new VisioChatManager(selectedModel, ApiConfig.GetWebhookUrl(), availableModels, libraryManager, AppendToChatHistory, aiChatPane);

                Debug.WriteLine("Starting webhook listener...");
                _ = StartWebhookListener(VISIO_PORT);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in ThisAddIn_Startup: {ex.Message}");
                MessageBox.Show($"Error during startup: {ex.Message}");
            }
        }

        private void AppendToChatHistory(string message)
        {
            // Make sure this method is invoked on the UI thread
            if (aiChatPane != null && !aiChatPane.IsDisposed)
            {
                if (aiChatPane.InvokeRequired)
                {
                    aiChatPane.Invoke(new Action<string>(AppendToChatHistory), message);
                }
                else
                {
                    aiChatPane.AppendToChatHistory(message);
                }
            }
        }

        // Ensure that SendShapesToN8nAsync is implemented as an async Task method in ThisAddIn.cs
        private async Task SendShapesToN8nAsync()
        {
            try
            {
                Debug.WriteLine("[SendShapesToN8nAsync] Preparing to send shape catalog to n8n...");
                var shapesCatalog = libraryManager.GetShapesCatalog();
                var jsonString = JsonConvert.SerializeObject(shapesCatalog);
                var content = new StringContent(jsonString, Encoding.UTF8, "application/json");

                string n8nWebhookUrl = ApiConfig.GetWebhookUrl("shape_catalog");

                var response = await httpClient.PostAsync(n8nWebhookUrl, content);
                response.EnsureSuccessStatusCode();

                Debug.WriteLine("[SendShapesToN8nAsync] Shape catalog sent successfully.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SendShapesToN8nAsync] Failed to send shape catalog: {ex.Message}");
            }
        }

        private async Task StartWebhookListener(int port)
        {
            listener = new HttpListener();
            
            // Add all necessary prefixes
            string[] endpoints = {
                "visio-command",
                "list-shapes",
                "image-agent",
                "chat-agent",
                "ShapeData"
            };

            // Always use VISIO_PORT (5680) for the listener
            Debug.WriteLine($"[StartWebhookListener] Using Visio port: {VISIO_PORT}");

            foreach (var endpoint in endpoints)
            {
                // Use direct paths for all endpoints
                var prefix = $"http://localhost:{VISIO_PORT}/{endpoint}/";
                listener.Prefixes.Add(prefix);
                Debug.WriteLine($"Added listener prefix: {prefix}");
            }

            try
            {
                listener.Start();
                Debug.WriteLine($"Webhook Listening on port {VISIO_PORT}");
                
                while (listener.IsListening)
                {
                    try 
                    {
                        HttpListenerContext context = await listener.GetContextAsync();
                        string requestPath = context.Request.Url.LocalPath;
                        string requestId = Guid.NewGuid().ToString();

                        Debug.WriteLine($"Received request on path: {requestPath}");

                        // Store the context for later use
                        activeConnections[requestId] = context;

                        if (requestPath == "/list-shapes/")
                        {
                            await HandleListShapesRequest(context);
                            CompleteRequest(requestId);
                        }
                        else if (requestPath == "/image-agent/")
                        {
                            string jsonResponse = await new System.IO.StreamReader(context.Request.InputStream).ReadToEndAsync();
                            Debug.WriteLine($"[Image-Agent] Received: {jsonResponse}");
                            AppendToChatHistory($"AI: {jsonResponse}");
                            CompleteRequest(requestId);
                        }
                        else if (requestPath == "/chat-agent/")
                        {
                            string jsonString = await new System.IO.StreamReader(context.Request.InputStream).ReadToEndAsync();
                            Debug.WriteLine($"[Chat-Agent] Received JSON: {jsonString}");
                            
                            try 
                            {
                                // Try to extract plain text from the JSON structure
                                string plainText = ExtractPlainText(jsonString);
                                Debug.WriteLine($"[Chat-Agent] Extracted plain text: {plainText}");
                                
                                // Format the text to remove duplicates and improve readability
                                string[] lines = plainText.Split(new[] { '\n', '.' }, StringSplitOptions.RemoveEmptyEntries);
                                var uniqueLines = new HashSet<string>(lines.Select(l => l.Trim()));
                                string formattedText = string.Join("\n", uniqueLines.Where(l => !string.IsNullOrWhiteSpace(l)));
                                
                                AppendToChatHistory($"AI: {formattedText}");
                                
                                // Only update the existing command if it exists
                                if (aiChatPane != null && !aiChatPane.IsDisposed)
                                {
                                    var currentCommand = aiChatPane.GetCurrentCommand();
                                    if (currentCommand != null)
                                    {
                                        currentCommand.Status = "Success";
                                        currentCommand.AIResponse = formattedText;
                                        aiChatPane.UpdateCommandStatus(currentCommand);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                // If parsing fails, try to clean up the raw string
                                Debug.WriteLine($"[Chat-Agent] Failed to parse JSON: {ex.Message}");
                                string cleanText = CleanupJsonText(jsonString);
                                AppendToChatHistory($"AI: {cleanText}");
                                
                                // Only update the existing command if it exists
                                if (aiChatPane != null && !aiChatPane.IsDisposed)
                                {
                                    var currentCommand = aiChatPane.GetCurrentCommand();
                                    if (currentCommand != null)
                                    {
                                        currentCommand.Status = "Failed";
                                        currentCommand.AIResponse = cleanText;
                                        aiChatPane.UpdateCommandStatus(currentCommand);
                                    }
                                }
                            }
                            
                            CompleteRequest(requestId);
                        }
                        else if (requestPath == "/visio-command/")
                        {
                            try
                            {
                                string jsonCommand = await new System.IO.StreamReader(context.Request.InputStream).ReadToEndAsync();
                                Debug.WriteLine($"[Visio-Command] Received command: {jsonCommand}");
                                
                                try
                                {
                                    await ProcessWebhookCommand(jsonCommand);
                                    await Task.Delay(200); // Increased delay to ensure command processing
                                    
                                    // Send success response
                                    var response = new { status = "success", message = "Command processed successfully" };
                                    var jsonResponse = JsonConvert.SerializeObject(response);
                                    var buffer = Encoding.UTF8.GetBytes(jsonResponse);
                                    
                                    context.Response.StatusCode = 200;
                                    context.Response.ContentType = "application/json";
                                    context.Response.ContentLength64 = buffer.Length;
                                    await context.Response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                                }
                                catch (Exception ex)
                                {
                                    // Send error response
                                    var response = new { status = "error", message = ex.Message };
                                    var jsonResponse = JsonConvert.SerializeObject(response);
                                    var buffer = Encoding.UTF8.GetBytes(jsonResponse);
                                    
                                    context.Response.StatusCode = 500;
                                    context.Response.ContentType = "application/json";
                                    context.Response.ContentLength64 = buffer.Length;
                                    await context.Response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                                }
                                finally
                                {
                                    context.Response.Close();
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[Visio-Command] Error processing request: {ex.Message}");
                                try
                                {
                                    var response = new { status = "error", message = "Internal server error" };
                                    var jsonResponse = JsonConvert.SerializeObject(response);
                                    var buffer = Encoding.UTF8.GetBytes(jsonResponse);
                                    
                                    context.Response.StatusCode = 500;
                                    context.Response.ContentType = "application/json";
                                    context.Response.ContentLength64 = buffer.Length;
                                    await context.Response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                                }
                                finally
                                {
                                    context.Response.Close();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error processing request: {ex.Message}");
                        continue; // Continue listening even if one request fails
                    }
                }
            }
            catch (HttpListenerException ex)
            {
                Debug.WriteLine($"[Error] Failed to start listener on port {port}: {ex.Message}");
            }
        }

        private void CompleteRequest(string requestId)
        {
            if (activeConnections.TryRemove(requestId, out HttpListenerContext context))
            {
                try
                {
                    byte[] buffer = Encoding.UTF8.GetBytes("Request processed.");
                    context.Response.ContentLength64 = buffer.Length;
                    context.Response.OutputStream.Write(buffer, 0, buffer.Length);
                    context.Response.Close();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[Error] Failed to complete request {requestId}: {ex.Message}");
                }
            }
        }

        private async Task HandleListShapesRequest(HttpListenerContext context)
        {
            var shapes = libraryManager.ListAllShapes();
            string jsonResponse = JsonConvert.SerializeObject(shapes);

            context.Response.ContentType = "application/json";
            byte[] buffer = Encoding.UTF8.GetBytes(jsonResponse);
            context.Response.ContentLength64 = buffer.Length;
            await context.Response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
        }

        public void StopWebhookListener()
        {
            if (listener != null)
            {
                listener.Stop();
                listener.Close();
                Debug.WriteLine("Webhook listener stopped.");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            StopWebhookListener();
        }

        private async Task ProcessWebhookCommand(string jsonCommand)
        {
            try
            {
                Debug.WriteLine($"[ProcessWebhookCommand] Received command: {jsonCommand}");

                if (commandProcessor == null)
                {
                    throw new InvalidOperationException("Command processor is not initialized.");
                }

                var affectedShapes = await Task.Run(() => commandProcessor.ProcessCommand(jsonCommand));
                Debug.WriteLine($"[ProcessWebhookCommand] Command processed. Received {affectedShapes?.Count ?? 0} shapes from processor");
                
                if (affectedShapes != null && affectedShapes.Any())
                {
                    Debug.WriteLine("[ProcessWebhookCommand] Shapes received:");
                    foreach (var shape in affectedShapes)
                    {
                        Debug.WriteLine($"[ProcessWebhookCommand] Shape: {shape}");
                    }

                    // Check if we can access AIChatPane
                    if (aiChatPane != null && !aiChatPane.IsDisposed)
                    {
                        var currentCommand = aiChatPane.GetCurrentCommand();
                        Debug.WriteLine($"[ProcessWebhookCommand] Current command found: {(currentCommand != null ? currentCommand.Id : "null")}");
                        if (currentCommand != null)
                        {
                            currentCommand.AffectedShapes = new List<ShapeInfo>(affectedShapes);
                            currentCommand.Status = "Success";
                            Debug.WriteLine($"[ProcessWebhookCommand] Added {affectedShapes.Count} shapes to command");
                            aiChatPane.UpdateCommandStatus(currentCommand);
                        }
                    }
                }
                else
                {
                    Debug.WriteLine("[ProcessWebhookCommand] No shapes were affected by the command");
                    if (aiChatPane != null && !aiChatPane.IsDisposed)
                    {
                        var currentCommand = aiChatPane.GetCurrentCommand();
                        if (currentCommand != null)
                        {
                            currentCommand.Status = "Failed";
                            currentCommand.AIResponse = "No shapes were created or modified.";
                            aiChatPane.UpdateCommandStatus(currentCommand);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessWebhookCommand] [Error] Failed to process webhook command: {ex.Message}");
                if (aiChatPane != null && !aiChatPane.IsDisposed)
                {
                    var currentCommand = aiChatPane.GetCurrentCommand();
                    if (currentCommand != null)
                    {
                        currentCommand.Status = "Failed";
                        currentCommand.AIResponse = $"Error: {ex.Message}";
                        aiChatPane.UpdateCommandStatus(currentCommand);
                    }
                }
                throw; // Re-throw to ensure the error is properly reported to the webhook caller
            }
        }

        public string[] GetCategories()
        {
            return libraryManager.GetCategories().ToArray();
        }

        public void OnRefreshLibrariesButtonClick(Office.IRibbonControl control)
        {
            try
            {
                libraryManager.LoadLibraries();
                if (Ribbon != null)
                {
                    Ribbon.InvalidateControl("CategorySelectionDropDown");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error while refreshing libraries: {ex.Message}");
            }
        }

        public int GetCategoryCount(Office.IRibbonControl control)
        {
            return libraryManager.GetCategories().Count();
        }

        public string GetCategoryLabel(Office.IRibbonControl control, int index)
        {
            return libraryManager.GetCategories().ElementAt(index);
        }

        public string GetSelectedCategoryID(Office.IRibbonControl control)
        {
            return CurrentCategory ?? string.Empty;
        }

        public void OnCategorySelectionChange(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var categories = libraryManager.GetCategories().ToArray();
            if (selectedIndex < 0 || selectedIndex >= categories.Length)
            {
                return;
            }

            // Set the current category in Globals
            CurrentCategory = selectedId;
            Debug.WriteLine($"[OnCategorySelectionChange] Current category set to: {CurrentCategory}");
        }

        public void OnAddTestShapeClick(Office.IRibbonControl control)
        {
            if (!string.IsNullOrEmpty(CurrentCategory))
            {
                var shapes = libraryManager.GetShapesInCategory(CurrentCategory).ToArray();
                if (shapes.Any())
                {
                    Random random = new Random();
                    string randomShape = shapes[random.Next(shapes.Length)];
                    var activePage = visioApplication.ActivePage;
                    double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                    double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                    // Calculate random position (in page units)
                    double randomX = random.NextDouble() * pageWidth;
                    double randomY = random.NextDouble() * pageHeight;

                    // Calculate a reasonable size for the shape (e.g., 5-10% of page width)
                    double minSize = Math.Min(pageWidth, pageHeight) * 0.05;
                    double maxSize = Math.Min(pageWidth, pageHeight) * 0.1;
                    double randomWidth = minSize + (random.NextDouble() * (maxSize - minSize));
                    double randomHeight = minSize + (random.NextDouble() * (maxSize - minSize));

                    // Convert to percentage of page size (as expected by AddShapeToDocument)
                    double xPercent = (randomX / pageWidth) * 100;
                    double yPercent = (randomY / pageHeight) * 100;
                    double widthPercent = (randomWidth / pageWidth) * 100;
                    double heightPercent = (randomHeight / pageHeight) * 100;

                    libraryManager.AddShapeToDocument(CurrentCategory, randomShape, xPercent, yPercent, widthPercent, heightPercent);

                    Debug.WriteLine($"Added random shape: {randomShape} at ({xPercent}%, {yPercent}%) with size ({widthPercent}%, {heightPercent}%)");
                }
            }
        }

        public void OnAPIEndpointChange(Office.IRibbonControl control, string text)
        {
            apiEndpoint = text;
        }

        public void OnConnectButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Debug.WriteLine("Starting connection to API and sending shape catalog...");

                // Run LoadModelsAsync and SendShapesToN8nAsync concurrently as background tasks
                Task.Run(async () =>
                {
                    await LoadModelsAsync();       // Load models from the API
                    await SendShapesToN8nAsync();  // Send shape catalog to n8n
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Unexpected error: {ex.Message}");
                MessageBox.Show($"Unexpected error: {ex.Message}");
            }
        }

        private async Task LoadModelsAsync()
        {
            try
            {
                Debug.WriteLine("Checking httpClient initialization...");
                if (httpClient == null) throw new NullReferenceException("httpClient is not initialized!");

                Debug.WriteLine("Checking apiEndpoint initialization...");
                if (string.IsNullOrEmpty(apiEndpoint)) throw new NullReferenceException("apiEndpoint is not initialized or is empty!");

                var requestBody = new
                {
                    command = "get_models"
                };

                libraryManager.LoadLibraries();
                if (Ribbon != null)
                {
                    Ribbon.InvalidateControl("CategorySelectionDropDown");
                }

                var jsonContent = new StringContent(JsonConvert.SerializeObject(requestBody), Encoding.UTF8, "application/json");

                var response = await httpClient.PostAsync($"{apiEndpoint}/connection_model_list", jsonContent);
                var responseContent = await response.Content.ReadAsStringAsync();

                Debug.WriteLine("Raw API Response: " + responseContent);

                List<string> modelList;
                try
                {
                    // First try parsing as a direct array
                    modelList = JsonConvert.DeserializeObject<List<string>>(responseContent);
                }
                catch
                {
                    try
                    {
                        // If that fails, try parsing as an object with a code property
                        var responseObj = JsonConvert.DeserializeObject<JObject>(responseContent);
                        if (responseObj["code"] != null && responseObj["code"].Type == JTokenType.Array)
                        {
                            modelList = responseObj["code"].ToObject<List<string>>();
                        }
                        else
                        {
                            throw new Exception("Unexpected response format");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error parsing response: {ex.Message}");
                        MessageBox.Show("Error parsing model list from API response.");
                        return;
                    }
                }

                Debug.WriteLine("Deserialized ModelResponse: " + (modelList?.Count ?? 0) + " models found.");

                if (modelList == null || !modelList.Any())
                {
                    Debug.WriteLine("Error: No models found.");
                    MessageBox.Show("No AI models available. Please check your Ollama installation.");
                    return;
                }

                uiControl.Invoke((MethodInvoker)(() =>
                {
                    Debug.WriteLine("Checking availableModels assignment...");
                    availableModels = modelList.ToArray();

                    Debug.WriteLine("Checking Ribbon initialization...");
                    if (Ribbon != null)
                    {
                        Ribbon.InvalidateControl("ConnectionStatus");
                        Ribbon.InvalidateControl("ModelSelectionDropDown");
                    }
                    else
                    {
                        Debug.WriteLine("Ribbon is null, skipping Ribbon invalidation.");
                    }

                    ShowAIChatPane();
                }));
            }
            catch (HttpRequestException httpEx)
            {
                Debug.WriteLine($"Error connecting to API: {httpEx.Message}");
                MessageBox.Show($"Error connecting to AI: {httpEx.Message}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Unexpected error: {ex.Message}");
                MessageBox.Show($"Unexpected error: {ex.Message}");
            }
        }

        public string GetModelLabel(Office.IRibbonControl control, int index)
        {
            if (availableModels != null && index >= 0 && index < availableModels.Length)
            {
                return availableModels[index];
            }
            return string.Empty;
        }

        public int GetModelCount(Office.IRibbonControl control)
        {
            return availableModels?.Length ?? 0;
        }

        private void ShowAIChatPane()
        {
            if (aiChatPane == null || aiChatPane.IsDisposed)
            {
                aiChatPane = new AIChatPane(selectedModel, apiEndpoint, availableModels, libraryManager);
                aiChatPane.FormClosed += (sender, e) => aiChatPane = null;

                IntPtr visioHandle = new IntPtr(visioApplication.WindowHandle32);
                if (visioHandle == IntPtr.Zero)
                {
                    aiChatPane.Show();
                }
                else
                {
                    aiChatPane.Show(new WindowWrapper(visioHandle));
                }
            }
            else
            {
                aiChatPane.BringToFront();
            }
        }

        public class WindowWrapper : IWin32Window
        {
            public WindowWrapper(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; }
        }

        private string ExtractPlainText(string jsonString)
        {
            try
            {
                // First try to parse as a simple JSON object with a text/message property
                var simpleObj = JsonConvert.DeserializeObject<dynamic>(jsonString);
                
                // Check for simple message or text property
                if (simpleObj.message != null) return simpleObj.message.ToString();
                if (simpleObj.text != null) return simpleObj.text.ToString();

                // If it's a more complex structure, try to find the first non-empty string value
                string plainText = FindFirstStringValue(simpleObj);
                if (!string.IsNullOrEmpty(plainText)) return plainText;

                // If we can't find a clear text value, convert the entire object to string
                // and clean it up
                string fullText = simpleObj.ToString();
                return CleanupJsonText(fullText);
            }
            catch
            {
                // If parsing fails, try to clean up the raw string
                return CleanupJsonText(jsonString);
            }
        }

        private string FindFirstStringValue(dynamic jsonObj)
        {
            if (jsonObj == null) return string.Empty;

            // If it's a simple string value
            if (jsonObj.Type == JTokenType.String)
                return jsonObj.Value;

            // If it's an object, search through its properties
            if (jsonObj.Type == JTokenType.Object)
            {
                foreach (var prop in jsonObj)
                {
                    string value = FindFirstStringValue(prop.Value);
                    if (!string.IsNullOrEmpty(value))
                        return value;
                }
            }

            // If it's an array, search through its elements
            if (jsonObj.Type == JTokenType.Array)
            {
                foreach (var item in jsonObj)
                {
                    string value = FindFirstStringValue(item);
                    if (!string.IsNullOrEmpty(value))
                        return value;
                }
            }

            return string.Empty;
        }

        private string CleanupJsonText(string text)
        {
            // First remove all JSON artifacts
            string cleaned = text
                .Replace("\"", "")
                .Replace("{", "")
                .Replace("}", "")
                .Replace("[", "")
                .Replace("]", "")
                .Replace("\\", "")
                .Replace(",", " ")
                .Replace(":", " ");

            // Split into lines and remove duplicates
            string[] lines = cleaned.Split(new[] { '\n', '.' }, StringSplitOptions.RemoveEmptyEntries);
            var uniqueLines = new HashSet<string>(
                lines.Select(l => l.Trim())
                    .Where(l => !string.IsNullOrWhiteSpace(l))
                    .Where(l => !l.StartsWith("n"))  // Remove numbered lines (n1, n2, etc.)
            );

            // Join unique lines with proper formatting
            string result = string.Join("\n", uniqueLines);

            // Clean up any remaining multiple spaces
            while (result.Contains("  "))
            {
                result = result.Replace("  ", " ");
            }

            return result.Trim();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}