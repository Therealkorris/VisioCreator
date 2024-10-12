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
using System.IO;
using System.Net.Http.Headers;

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
            if (string.IsNullOrEmpty(userMessage)) return;

            AppendToChatHistory("User: " + userMessage);
            chatInput.Clear();

            try
            {
                // Prepare the form content to match what FastAPI expects
                var content = new MultipartFormDataContent();
                content.Add(new StringContent(userMessage), "prompt");
                content.Add(new StringContent(selectedModel), "model");

                // Send the request to FastAPI server
                var response = await httpClient.PostAsync($"{pythonApiEndpoint}/text-prompt", content);

                // Log request data for debugging
                Debug.WriteLine($"Sent message: {userMessage}, model: {selectedModel}");

                // Process the response stream
                var responseStream = await response.Content.ReadAsStreamAsync();
                using (var reader = new StreamReader(responseStream))
                {
                    string line;
                    StringBuilder fullResponse = new StringBuilder();

                    // Read the response and accumulate the full response
                    while ((line = await reader.ReadLineAsync()) != null)
                    {
                        fullResponse.Append(line);
                    }

                    // Append the AI response to chat history
                    AppendToChatHistory("AI: " + fullResponse.ToString().Trim());
                    Debug.WriteLine($"AI Response: {fullResponse.ToString().Trim()}");
                }
            }
            catch (Exception ex)
            {
                AppendToChatHistory("Error: " + ex.Message);
                Debug.WriteLine("Error sending message: " + ex.Message);
            }
        }

        private async Task SendMessageWithImage(string imagePath)
        {
            string userMessage = chatInput.Text.Trim();
            if (string.IsNullOrEmpty(userMessage)) return;

            AppendToChatHistory("User: " + userMessage);
            chatInput.Clear();

            try
            {
                var content = new MultipartFormDataContent();
                content.Add(new StringContent(userMessage), "prompt");
                content.Add(new StringContent(selectedModel), "model");

                // Attach the image if provided
                if (!string.IsNullOrEmpty(imagePath))
                {
                    var imageContent = new ByteArrayContent(File.ReadAllBytes(imagePath));
                    imageContent.Headers.ContentType = new MediaTypeHeaderValue("image/jpeg"); // or "image/png"
                    content.Add(imageContent, "file", Path.GetFileName(imagePath));
                }

                // Send the request to FastAPI server
                var response = await httpClient.PostAsync($"{pythonApiEndpoint}/image-prompt", content);

                // Log the request and response for debugging
                Debug.WriteLine($"Sent prompt: {userMessage}, model: {selectedModel}, image: {Path.GetFileName(imagePath)}");

                // Read the response from FastAPI
                var responseStream = await response.Content.ReadAsStreamAsync();
                using (var reader = new StreamReader(responseStream))
                {
                    string line;
                    StringBuilder fullResponse = new StringBuilder();
                    while ((line = await reader.ReadLineAsync()) != null)
                    {
                        fullResponse.Append(line);
                    }
                    AppendToChatHistory("AI: " + fullResponse.ToString().Trim());
                    Debug.WriteLine($"AI Response: {fullResponse.ToString().Trim()}");
                }
            }
            catch (Exception ex)
            {
                AppendToChatHistory("Error: " + ex.Message);
                Debug.WriteLine("Error sending message: " + ex.Message);
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
                AppendToChatHistory("System: Unknown command received.");
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
                    AppendToChatHistory($"System: Created shape '{shape}' from category '{category}' at position ({x}, {y})");
                    return;
                }
            }

            AppendToChatHistory($"System: No shapes found matching '{shapeType}'.");
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
            Debug.WriteLine($"Updated available models in dropdown: {string.Join(", ", models)}");
        }

        public class ModelResponse
        {
            public List<string> Models { get; set; }
        }
    }
}
