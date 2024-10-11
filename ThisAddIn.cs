using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;
using Microsoft.Office.Tools.Ribbon;
using OllamaSharp;  // OllamaSharp namespace for Ollama API
using System.Windows.Forms;

namespace VisioPlugin
{
    [ComVisible(true)]
    public partial class ThisAddIn
    {
        private Visio.Application visioApplication;
        internal Office.IRibbonUI Ribbon { get; set; }
        private LibraryManager libraryManager;
        internal string CurrentCategory { get; set; }
        private string apiEndpoint = "http://localhost:11434";  // Default API endpoint
        public bool isConnected = false;  // Connection status
        private string[] availableModels = new string[0];  // Placeholder for AI models
        private OllamaApiClient ollamaClient;  // OllamaSharp API client
        private string selectedModel = "llama3.1:8b";  // Default model for the API

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonExtension(this);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            visioApplication = (Visio.Application)Application;
            libraryManager = new LibraryManager(visioApplication);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        // Method to get the available categories
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
                    Debug.WriteLine("Invalidating ribbon...");
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
                Debug.WriteLine($"Invalid category index: {selectedIndex}");
                return;
            }

            CurrentCategory = selectedId;
            Debug.WriteLine($"Current category is now set to: {CurrentCategory}");
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
                    double randomX = random.NextDouble() * activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                    double randomY = random.NextDouble() * activePage.PageSheet.CellsU["PageHeight"].ResultIU;
                    libraryManager.AddShapeToDocument(CurrentCategory, randomShape, randomX, randomY);
                }
            }
        }

        // Handles AI API endpoint changes
        public void OnAPIEndpointChange(Office.IRibbonControl control, string text)
        {
            apiEndpoint = text;
        }

        // Attempts to connect to the AI API using OllamaSharp
        public async void OnConnectButtonClick(Office.IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("OnConnectButtonClick method started");
            Debug.WriteLine("OnConnectButtonClick method started");
            try
            {
                Debug.WriteLine($"Connecting to Ollama API at {apiEndpoint}...");

                // Initialize OllamaSharp API client
                var uri = new Uri(apiEndpoint);
                ollamaClient = new OllamaApiClient(uri);

                Debug.WriteLine("OllamaApiClient initialized. Attempting to list models...");

                // Check available models
                var models = await ollamaClient.ListLocalModels();
                Debug.WriteLine($"ListLocalModels call completed. Response: {models}");

                if (models != null && models.Any())
                {
                    isConnected = true;
                    Debug.WriteLine("Successfully connected to Ollama API and retrieved models.");
                    availableModels = models.Select(m => m.Name).ToArray();
                    Debug.WriteLine($"Available models: {string.Join(", ", availableModels)}");
                }
                else
                {
                    isConnected = false;
                    Debug.WriteLine("No models available in Ollama or models list is null.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error connecting to Ollama API: {ex.Message}");
                Debug.WriteLine($"Stack Trace: {ex.StackTrace}");
                isConnected = false;
            }

            Debug.WriteLine($"Connection status: {(isConnected ? "Connected" : "Not Connected")}");
            Debug.WriteLine($"Number of available models: {availableModels?.Length ?? 0}");

            // Update the ribbon controls
            Debug.WriteLine("Attempting to invalidate ribbon controls...");
            try
            {
                Ribbon?.InvalidateControl("ConnectionStatus");
                Ribbon?.InvalidateControl("ModelSelectionDropDown");
                Debug.WriteLine("Ribbon controls invalidated successfully.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error invalidating ribbon controls: {ex.Message}");
            }

            Debug.WriteLine("OnConnectButtonClick method completed");
        }

        // Provides the label for each model in the AI model selection dropdown
        // Provides the label for each model in the AI model selection dropdown
        public string GetModelLabel(Office.IRibbonControl control, int index)
        {
            Debug.WriteLine($"GetModelLabel called for index {index}");
            if (availableModels != null && index >= 0 && index < availableModels.Length)
            {
                Debug.WriteLine($"Returning model name: {availableModels[index]}");
                return availableModels[index];
            }
            Debug.WriteLine("Returning empty string");
            return string.Empty;
        }

        // Provides the number of available AI models
        public int GetModelCount(Office.IRibbonControl control)
        {
            Debug.WriteLine($"GetModelCount called. Returning {availableModels?.Length ?? 0}");
            return availableModels?.Length ?? 0;
        }

        // Handles the event when an AI model is selected
        public void OnModelSelectionChange(Office.IRibbonControl control, string selectedItemId)
        {
            Debug.WriteLine($"Model selected: {selectedItemId}");
            selectedModel = selectedItemId;  // Set the selected model
        }

        // Loads the available AI models asynchronously using OllamaSharp
        private async Task LoadAvailableModels()
        {
            try
            {
                // Fetch models from Ollama (this is already a list of Model objects)
                var models = await ollamaClient.ListLocalModels();

                // Ensure models exist
                if (models != null && models.Any())
                {
                    Debug.WriteLine("Ollama Models Retrieved:");
                    foreach (var model in models)
                    {
                        Debug.WriteLine($"Model Name: {model.Name}, Size: {model.Size}, Modified: {model.ModifiedAt}");
                    }

                    // Convert model names to a string array for the dropdown
                    availableModels = models.Select(model => model.Name).ToArray();

                    // Invalidate the dropdown control to refresh the model list in the UI
                    Ribbon?.InvalidateControl("ModelSelectionDropDown");
                }
                else
                {
                    Debug.WriteLine("No models available in Ollama.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading models from Ollama: {ex.Message}");
                availableModels = new string[0];
            }
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
