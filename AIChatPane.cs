using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using Newtonsoft.Json;
using Visio = Microsoft.Office.Interop.Visio;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Collections.Generic;

namespace VisioPlugin
{
    public partial class AIChatPane : Form
    {
        private readonly Visio.Application visioApplication;
        private readonly LibraryManager libraryManager;
        private string selectedModel;
        private readonly string pythonApiEndpoint;

        private TextBox chatInput;
        private Button sendButton;
        private RichTextBox chatHistory;
        private ComboBox modelDropdown;
        private Label modelLabel;
        private string[] availableModels;
        private static readonly HttpClient httpClient = new HttpClient();

        public AIChatPane(Visio.Application visioApp, LibraryManager libManager, string model, string apiEndpoint, string[] availableModels)
        {
            visioApplication = visioApp;
            libraryManager = libManager;
            selectedModel = model;
            pythonApiEndpoint = apiEndpoint;
            this.availableModels = availableModels;
            InitializeComponent();
            PopulateModelDropdown();
        }

        private async Task LoadAvailableModelsAsync()
        {
            try
            {
                var response = await httpClient.GetAsync($"{pythonApiEndpoint}/models");
                response.EnsureSuccessStatusCode();
                var responseContent = await response.Content.ReadAsStringAsync();
                var modelResponse = JsonConvert.DeserializeObject<ModelResponse>(responseContent);
                availableModels = modelResponse?.Models?.ToArray() ?? Array.Empty<string>();
                
                // Ensure UI updates happen on the UI thread
                this.Invoke((MethodInvoker)delegate 
                {
                    PopulateModelDropdown();
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading models: {ex.Message}");
                this.Invoke((MethodInvoker)delegate 
                {
                    MessageBox.Show("Error loading models. Please try again later.");
                });
            }
        }

        private async Task LoadAvailableModels()
        {
        try
        {
            var response = await httpClient.GetAsync($"{pythonApiEndpoint}/models");
            response.EnsureSuccessStatusCode();
            var responseContent = await response.Content.ReadAsStringAsync();
            var modelResponse = JsonConvert.DeserializeObject<ModelResponse>(responseContent);
            availableModels = modelResponse?.Models?.ToArray() ?? Array.Empty<string>();
            PopulateModelDropdown();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error loading models: {ex.Message}");
            MessageBox.Show("Error loading models. Please try again later.");
        }
    }

        private void PopulateModelDropdown()
        {
            if (availableModels != null && availableModels.Any())
            {
                modelDropdown.Items.Clear();
                modelDropdown.Items.AddRange(availableModels);
                modelDropdown.SelectedItem = selectedModel;
            }
            else
            {
                MessageBox.Show("No models available.");
            }
        }

        private void InitializeComponent()
        {
            chatHistory = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.LightYellow,
                Font = new Font("Arial", 10),
                BorderStyle = BorderStyle.None
            };

            chatInput = new TextBox
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                Multiline = true,
                Font = new Font("Arial", 10),
                BorderStyle = BorderStyle.FixedSingle
            };
            chatInput.KeyDown += ChatInput_KeyDown;

            sendButton = new Button
            {
                Text = "Send",
                Dock = DockStyle.Bottom,
                Height = 40,
                BackColor = Color.SteelBlue,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Arial", 10, FontStyle.Bold)
            };
            sendButton.Click += SendButton_Click;

            modelLabel = new Label
            {
                Text = "Select AI Model:",
                Dock = DockStyle.Top,
                Height = 20,
                Font = new Font("Arial", 10, FontStyle.Bold),
                ForeColor = Color.SteelBlue
            };

            modelDropdown = new ComboBox
            {
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Font("Arial", 10),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            modelDropdown.SelectedIndexChanged += ModelDropdown_SelectedIndexChanged;

            Controls.Add(chatHistory);
            Controls.Add(chatInput);
            Controls.Add(sendButton);
            Controls.Add(modelDropdown);
            Controls.Add(modelLabel);

            this.Layout += (sender, e) => PerformLayout();
        }

        private void ModelDropdown_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedModel = modelDropdown.SelectedItem?.ToString();
            Debug.WriteLine($"Selected model updated to: {selectedModel}");
        }

        private async void ChatInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !e.Shift)
            {
                e.SuppressKeyPress = true;
                await SendMessage();
            }
        }

        private async void SendButton_Click(object sender, EventArgs e)
        {
            await SendMessage();
        }

        private async Task SendMessage()
        {
            string userMessage = chatInput.Text.Trim();
            if (string.IsNullOrEmpty(userMessage))
                return;

            AppendToChatHistory("User: " + userMessage);
            chatInput.Clear();

            try
            {
                var payload = new
                {
                    prompt = userMessage,
                    model = selectedModel
                };

                var jsonContent = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
                var response = await httpClient.PostAsync($"{pythonApiEndpoint}/execute-command", jsonContent);
                var responseContent = await response.Content.ReadAsStringAsync();

                Debug.WriteLine($"AI Response: {responseContent}");

                var aiResponse = JsonConvert.DeserializeObject<Dictionary<string, string>>(responseContent);
                if (aiResponse.ContainsKey("response"))
                {
                    AppendToChatHistory("System: " + aiResponse["response"]);
                }
                else if (aiResponse.ContainsKey("error"))
                {
                    AppendToChatHistory("Error: " + aiResponse["error"]);
                }
            }
            catch (Exception ex)
            {
                AppendToChatHistory("Error: " + ex.Message);
            }
        }

        private void AppendToChatHistory(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(AppendToChatHistory), message);
                return;
            }

            chatHistory.AppendText(message + Environment.NewLine);
            chatHistory.ScrollToCaret();
        }

        // Communicate with the FastAPI server running locally on Python
        private async Task<dynamic> GetAIResponseFromServer(string message)
        {
            try
            {
                var requestData = new
                {
                    command = message,
                    // Corrected the syntax issue here by adding a comma
                    @params = new { model = selectedModel } // `params` is a reserved keyword, so it's escaped with `@`
                };

                var jsonContent = new StringContent(JsonConvert.SerializeObject(requestData), Encoding.UTF8, "application/json");
                Debug.WriteLine($"Sending Payload: {JsonConvert.SerializeObject(requestData)}");
                var response = await httpClient.PostAsync("http://127.0.0.1:8000/execute-command", jsonContent);
                response.EnsureSuccessStatusCode();

                string responseBody = await response.Content.ReadAsStringAsync();
                return JsonConvert.DeserializeObject<dynamic>(responseBody);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in GetAIResponseFromServer: {ex.Message}");
                return null;
            }
        }

        private void ExecuteVisioCommand(dynamic commandResponse)
        {
            if (commandResponse?.result?.action == "create_shape")
            {
                string shape = commandResponse.result.shape;
                double x = commandResponse.result.position.x;
                double y = commandResponse.result.position.y;

                CreateShapeFromLibrary(shape, x, y);
            }
            else
            {
                AppendMessage("System", "Unknown command received.");
            }
        }

        private void CreateShapeFromLibrary(string shapeType, double x, double y)
        {
            var categories = libraryManager.GetCategories();
            foreach (var category in categories)
            {
                var shapes = libraryManager.GetShapesInCategory(category);
                var shape = shapes.FirstOrDefault(s => s.ToLower().Contains(shapeType));
                if (shape != null)
                {
                    libraryManager.AddShapeToDocument(category, shape, x, y);
                    AppendMessage("System", $"Created shape '{shape}' from category '{category}' at position ({x}, {y})");
                    return;
                }
            }

            AppendMessage("System", $"No shapes found matching '{shapeType}'.");
        }

        private void AppendMessage(string sender, string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => AppendMessage(sender, message)));
            }
            else
            {
                chatHistory.AppendText($"{sender}: {message}\n\n");
                chatHistory.ScrollToCaret();
            }
        }

        public void UpdateAvailableModels(string[] models)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string[]>(UpdateAvailableModels), new object[] { models });
                return;
            }

            modelDropdown.Items.Clear();
            modelDropdown.Items.AddRange(models);
            if (!string.IsNullOrEmpty(selectedModel) && models.Contains(selectedModel))
            {
                modelDropdown.SelectedItem = selectedModel;
            }
            else if (models.Length > 0)
            {
                modelDropdown.SelectedIndex = 0;
            }
        }
    }

    public class ModelResponse
    {
        public List<string> Models { get; set; }
    }
}
