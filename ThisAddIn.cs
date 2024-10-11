// ThisAddIn.cs
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;
using Microsoft.Office.Tools.Ribbon;
using OllamaSharp;
using OllamaSharp.Models;
using System.Windows.Forms;
using System.Threading;
using System.Net.Http;

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
        private string apiEndpoint = "http://localhost:11434";
        public bool isConnected = false;
        private string[] availableModels = new string[0];
        private OllamaApiClient ollamaClient;
        private string selectedModel = "llama3.1:8b";
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

        public async void OnConnectButtonClick(Office.IRibbonControl control)
        {
            try
            {
                var uri = new Uri(apiEndpoint);
                ollamaClient = new OllamaApiClient(uri);

                var models = await ollamaClient.ListLocalModels();

                uiControl.Invoke((MethodInvoker)(() =>
                {
                    if (models != null && models.Any())
                    {
                        isConnected = true;
                        availableModels = models.Select(m => m.Name).ToArray();
                        ShowAIChatPane();
                    }
                    else
                    {
                        isConnected = false;
                        MessageBox.Show("No AI models available. Please check your Ollama installation.");
                    }

                    Ribbon?.InvalidateControl("ConnectionStatus");
                    Ribbon?.InvalidateControl("ModelSelectionDropDown");
                }));
            }
            catch (HttpRequestException httpEx)
            {
                Debug.WriteLine($"Error connecting to Ollama API: {httpEx.Message}");
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

        public void OnModelSelectionChange(Office.IRibbonControl control, string selectedItemId)
        {
            Debug.WriteLine($"Model selected: {selectedItemId}");
            selectedModel = selectedItemId;  // Store the selected model
        }

        private async Task LoadAvailableModels()
        {
            try
            {
                var models = await ollamaClient.ListLocalModels();

                if (models != null && models.Any())
                {
                    availableModels = models.Select(model => model.Name).ToArray();
                    Ribbon?.InvalidateControl("ModelSelectionDropDown");
                }
            }
            catch (Exception ex)
            {
                availableModels = new string[0];
            }
        }

        private void ShowAIChatPane()
        {
            if (aiChatPane == null || aiChatPane.IsDisposed)
            {
                aiChatPane = new AIChatPane(visioApplication, ollamaClient, selectedModel);  // Pass selectedModel
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
