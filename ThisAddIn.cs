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
using System.Net;

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

        public string apiEndpoint = "http://localhost:5678/webhook";
        public bool isConnected = false;
        private string[] availableModels = new string[0];
        private HttpClient httpClient = new HttpClient();
        private string selectedModel = "llama3.2";
        private AIChatPane aiChatPane;
        private VisioCommandProcessor commandProcessor;
        private HttpListener listener;
        private VisioChatManager visioChatManager;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonExtension(this);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Debug.WriteLine("Initializing Visio application...");
                visioApplication = (Visio.Application)Application;

                Debug.WriteLine("Initializing LibraryManager...");
                libraryManager = new LibraryManager(visioApplication);

                Debug.WriteLine("Initializing UIControl...");
                uiControl = new System.Windows.Forms.Control();
                uiControl.CreateControl();

                commandProcessor = new VisioCommandProcessor(visioApplication, libraryManager);

                Debug.WriteLine("Initializing VisioChatManager...");
                visioChatManager = new VisioChatManager(selectedModel, apiEndpoint, availableModels, libraryManager, AppendToChatHistory, aiChatPane);

                Debug.WriteLine("Starting webhook listener...");
                _ = StartWebhookListener(5680);
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

                string n8nWebhookUrl = "http://localhost:5678/webhook/shape_catalog";

                var response = await httpClient.PostAsync(n8nWebhookUrl, content);
                response.EnsureSuccessStatusCode();

                Debug.WriteLine("[SendShapesToN8nAsync] Shape catalog sent successfully to /webhook/shape_catalog.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SendShapesToN8nAsync] Failed to send shape catalog: {ex.Message}");
            }
        }

        private async Task StartWebhookListener(int port)
        {
            listener = new HttpListener();
            listener.Prefixes.Add($"http://localhost:{port}/visio-command/");
            listener.Prefixes.Add($"http://localhost:{port}/list-shapes/");
            listener.Prefixes.Add($"http://localhost:{port}/image-agent/"); // Changed path
            try
            {
                listener.Start();
                Debug.WriteLine($"Webhook Listening for Visio commands on port {port}");
                while (listener.IsListening)
                {
                    HttpListenerContext context = await listener.GetContextAsync();
                    string requestPath = context.Request.Url.LocalPath;
                    if (requestPath == "/list-shapes/")
                    {
                        await HandleListShapesRequest(context);
                    }
                    else if (requestPath == "/image-agent/") // Changed path
                    {
                        //string jsonResponse = await new System.IO.StreamReader(context.Request.InputStream).ReadToEndAsync();
                        //AppendToChatHistory($"AI: {jsonResponse}");
                    }
                    else
                    {
                        string jsonCommand = await new System.IO.StreamReader(context.Request.InputStream).ReadToEndAsync();
                        await ProcessWebhookCommand(jsonCommand);
                    }

                    HttpListenerResponse response = context.Response;
                    byte[] buffer = Encoding.UTF8.GetBytes("Request processed.");
                    response.ContentLength64 = buffer.Length;
                    await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                    response.OutputStream.Close();
                }
            }
            catch (HttpListenerException ex)
            {
                Debug.WriteLine($"[Error] Failed to start listener on port {port}: {ex.Message}");
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

                if (commandProcessor != null)
                {
                    await Task.Run(() => commandProcessor.ProcessCommand(jsonCommand));
                    Debug.WriteLine("[ProcessWebhookCommand] Command forwarded to VisioCommandProcessor.");
                }
                else
                {
                    Debug.WriteLine("[ProcessWebhookCommand] [Error] CommandProcessor is not initialized.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessWebhookCommand] [Error] Failed to process webhook command: {ex.Message}");
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

                var modelList = JsonConvert.DeserializeObject<List<string>>(responseContent);

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

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}