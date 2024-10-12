﻿using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Collections.Generic;
using Newtonsoft.Json;

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
        private string apiEndpoint = "http://localhost:11434"; // API for Ollama
        private string pythonApiEndpoint = "http://localhost:8000"; // Python FastAPI endpoint
        public bool isConnected = false;
        private string[] availableModels = new string[0];
        private HttpClient httpClient = new HttpClient();
        private string selectedModel = "llama3.1:8b"; // Default model
        private AIChatPane aiChatPane;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonExtension(this);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            visioApplication = (Visio.Application)Application;
            libraryManager = new LibraryManager(visioApplication);
            uiControl = new System.Windows.Forms.Control();
            uiControl.CreateControl();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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

            CurrentCategory = selectedId;
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

        public void OnAPIEndpointChange(Office.IRibbonControl control, string text)
        {
            apiEndpoint = text;
        }

        public void OnConnectButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Debug.WriteLine("Starting connection to API...");
                // Trigger the task without blocking the UI
                Task.Run(async () => await LoadModelsAsync()); // Background async task
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
                var uri = new Uri(apiEndpoint);
                var response = await httpClient.GetAsync($"{pythonApiEndpoint}/models");
                var responseContent = await response.Content.ReadAsStringAsync();

                // Log the raw response
                Debug.WriteLine("Raw API Response: " + responseContent);

                // Deserialize using Newtonsoft.Json
                var modelResponse = JsonConvert.DeserializeObject<ModelResponse>(responseContent);

                // Log the deserialized object
                Debug.WriteLine("Deserialized ModelResponse: " + (modelResponse?.Models?.Count ?? 0) + " models found.");

                // Ensure the response contains models
                if (modelResponse == null || modelResponse.Models == null || !modelResponse.Models.Any())
                {
                    Debug.WriteLine("Error: No models found.");
                    MessageBox.Show("No AI models available. Please check your Ollama installation.");
                    return;
                }

                // UI update must happen on the main thread (InvokeRequired pattern)
                uiControl.Invoke((MethodInvoker)(() =>
                {
                    isConnected = true;
                    availableModels = modelResponse.Models.ToArray(); // Store available models
                    Debug.WriteLine("Models loaded successfully.");

                    // Invalidate and update Ribbon to show the model dropdown
                    Ribbon?.InvalidateControl("ConnectionStatus");
                    Ribbon?.InvalidateControl("ModelSelectionDropDown");

                    ShowAIChatPane(); // Load the AI chat window
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

        // Helper class to deserialize the models response
        public class ModelResponse
        {
            public List<string> Models { get; set; }
        }

        public string GetModelLabel(Office.IRibbonControl control, int index)
        {
            // Ensure the index is within bounds and there are available models
            if (availableModels != null && index >= 0 && index < availableModels.Length)
            {
                return availableModels[index];
            }
            return string.Empty;
        }

        public int GetModelCount(Office.IRibbonControl control)
        {
            return availableModels?.Length ?? 0;  // Return the number of available models
        }

        public async void OnModelSelectionChange(Office.IRibbonControl control, string selectedItemId)
        {
            Debug.WriteLine($"Model selected: {selectedItemId}");
            selectedModel = selectedItemId;  // Store the selected model

            // Send model selection to the Python backend
            await SendModelSelectionToPython(selectedModel);
        }

        private async Task SendModelSelectionToPython(string model)
        {
            try
            {
                var modelSelectionPayload = new { model = model };
                // Use Newtonsoft.Json to serialize the payload
                var jsonContent = new StringContent(JsonConvert.SerializeObject(modelSelectionPayload), Encoding.UTF8, "application/json");

                var response = await httpClient.PostAsync($"{pythonApiEndpoint}/set-model", jsonContent);
                if (response.IsSuccessStatusCode)
                {
                    Debug.WriteLine("Model successfully updated on Python backend.");
                }
                else
                {
                    Debug.WriteLine($"Error updating model on Python backend: {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error communicating with Python backend: {ex.Message}");
            }
        }

        private void ShowAIChatPane()
        {
            if (aiChatPane == null || aiChatPane.IsDisposed)
            {
                aiChatPane = new AIChatPane(selectedModel, pythonApiEndpoint, availableModels);
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
