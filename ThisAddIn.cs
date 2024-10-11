using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;
using Microsoft.Office.Tools.Ribbon;
using OllamaSharp;  // OllamaSharp namespace for Ollama API

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
            try
            {
                Debug.WriteLine($"Connecting to Ollama API at {apiEndpoint}...");

                // Initialize OllamaSharp API client
                var uri = new Uri(apiEndpoint);
                ollamaClient = new OllamaApiClient(uri);

                // Check available models
                var models = await ollamaClient.ListLocalModels();
                if (models.Any())
                {
                    isConnected = true;
                    Debug.WriteLine("Successfully connected to Ollama API.");
                }
                else
                {
                    isConnected = false;
                    Debug.WriteLine("No models available in Ollama.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error connecting to Ollama API: {ex.Message}");
                isConnected = false;
            }

            // Update the ribbon control for connection status
            Ribbon?.InvalidateControl("ConnectionStatus");
        }

        // Provides the label for each model in the AI model selection dropdown
        public string GetModelLabel(Office.IRibbonControl control, int index)
        {
            return availableModels[index];
        }

        // Provides the number of available AI models
        public int GetModelCount(Office.IRibbonControl control)
        {
            return availableModels.Length;
        }

        // Handles the event when an AI model is selected
        public void OnModelSelectionChange(Office.IRibbonControl control, string selectedItemId)
        {
            Debug.WriteLine($"Model selected: {selectedItemId}");
            selectedModel = selectedItemId;
        }

        // Loads the available AI models asynchronously using OllamaSharp
        private async Task LoadAvailableModels()
        {
            try
            {
                // Fetch the models and convert them to strings (assuming 'Name' is the property you need)
                var models = await ollamaClient.ListLocalModels();
                availableModels = models.Select(model => model.Name).ToArray();  // Convert to string array
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading models: {ex.Message}");
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
