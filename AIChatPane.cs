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
        private TextBox chatInput;
        private Button sendButton;
        private RichTextBox chatHistory;
        private ComboBox modelDropdown;
        private Label modelLabel;
        private Visio.Application visioApplication;
        private string selectedModel;
        private string[] availableModels;
        private LibraryManager libraryManager;
        private static readonly HttpClient httpClient = new HttpClient();

        public AIChatPane(Visio.Application visioApp, string initialModel)
        {
            visioApplication = visioApp;
            this.selectedModel = initialModel;
            libraryManager = new LibraryManager(visioApp);
            libraryManager.LoadLibraries();
            InitializeComponent();

            // Load models asynchronously, without blocking the UI
            _ = LoadAvailableModelsAsync(); // Fire and forget, UI continues to load
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

        private async Task LoadAvailableModelsAsync()
        {
            try
            {
                HttpResponseMessage response = await httpClient.GetAsync("http://127.0.0.1:8000/models");
                response.EnsureSuccessStatusCode();

                string responseBody = await response.Content.ReadAsStringAsync();

                // Parse the response using a strongly typed class
                var modelsResponse = JsonConvert.DeserializeObject<ModelsResponse>(responseBody);

                if (modelsResponse != null && modelsResponse.Models.Any())
                {
                    // Populate the dropdown with models
                    Invoke(new Action(() => {
                        modelDropdown.Items.Clear();
                        modelDropdown.Items.AddRange(modelsResponse.Models.ToArray());
                        modelDropdown.SelectedItem = selectedModel; // Set the default
                    }));
                }
                else
                {
                    MessageBox.Show("No models available.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading models: {ex.Message}");
            }
        }


        // Define the class to match the API response
        public class ModelsResponse
        {
            public List<string> Models { get; set; }
        }


        private void ModelDropdown_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedModel = modelDropdown.SelectedItem?.ToString();
            Debug.WriteLine($"Selected model updated to: {selectedModel}");
        }

        private void ChatInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !e.Shift)
            {
                e.SuppressKeyPress = true;
                SendMessage();
            }
        }

        private void SendButton_Click(object sender, EventArgs e)
        {
            SendMessage();
        }

        private async void SendMessage()
        {
            string userMessage = chatInput.Text.Trim();
            if (!string.IsNullOrEmpty(userMessage))
            {
                AppendMessage("User", userMessage);
                chatInput.Clear();
                sendButton.Enabled = false;

                // Send message to AI and get the response
                var aiResponse = await GetAIResponseFromServer(userMessage);

                if (aiResponse != null)
                {
                    Debug.WriteLine("AI Response: " + JsonConvert.SerializeObject(aiResponse));  // Debugging info
                    // Parse response and handle command execution
                    ExecuteVisioCommand(aiResponse);
                }
                else
                {
                    AppendMessage("System", "No response from AI.");
                }

                sendButton.Enabled = true;
            }
        }

        private async Task<dynamic> GetAIResponseFromServer(string message)
        {
            try
            {
                var requestData = new
                {
                    command = message,
                    @params = new { model = selectedModel }
                };

                var jsonContent = new StringContent(JsonConvert.SerializeObject(requestData), Encoding.UTF8, "application/json");
                var response = await httpClient.PostAsync("http://127.0.0.1:8000/execute-command", jsonContent);
                response.EnsureSuccessStatusCode();

                string responseBody = await response.Content.ReadAsStringAsync();
                return JsonConvert.DeserializeObject<dynamic>(responseBody);
            }
            catch (Exception ex)
            {
                AppendMessage("System", $"Error: {ex.Message}");
                Debug.WriteLine($"Error in GetAIResponseFromServer: {ex.Message}");  // Debugging info
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
    }
}
