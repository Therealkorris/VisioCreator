using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using OllamaSharp;
using OllamaSharp.Models;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;

namespace VisioPlugin
{
    public class AIChatPane : Form
    {
        private TextBox chatInput;
        private Button sendButton;
        private RichTextBox chatHistory;
        private ComboBox modelDropdown;
        private Label modelLabel;
        private Visio.Application visioApplication;
        private OllamaApiClient ollamaClient;
        private string selectedModel;
        private string[] availableModels;

        public AIChatPane(Visio.Application visioApp, OllamaApiClient ollamaClient, string initialModel)
        {
            visioApplication = visioApp;
            this.ollamaClient = ollamaClient;
            this.selectedModel = initialModel;
            InitializeComponent();
            LoadAvailableModels();  // Load the available models into the dropdown

            this.Text = "AI Chat";
            this.Size = new System.Drawing.Size(400, 600);
        }

        private void InitializeComponent()
        {
            // Initialize chat history
            chatHistory = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true
            };

            // Initialize chat input
            chatInput = new TextBox
            {
                Dock = DockStyle.Bottom,
                Height = 50
            };

            // Initialize send button
            sendButton = new Button
            {
                Text = "Send",
                Dock = DockStyle.Bottom,
                Height = 30
            };
            sendButton.Click += SendButton_Click;

            // Initialize model label
            modelLabel = new Label
            {
                Text = "Select AI Model:",
                Dock = DockStyle.Top,
                Height = 20
            };

            // Initialize model dropdown
            modelDropdown = new ComboBox
            {
                Dock = DockStyle.Top,
                Height = 30
            };
            modelDropdown.SelectedIndexChanged += ModelDropdown_SelectedIndexChanged;

            // Add controls to the form
            Controls.Add(chatHistory);
            Controls.Add(chatInput);
            Controls.Add(sendButton);
            Controls.Add(modelDropdown);
            Controls.Add(modelLabel);

            this.Layout += (sender, e) => PerformLayout();
        }

        // Load available models and populate the dropdown
        private async void LoadAvailableModels()
        {
            try
            {
                var models = await ollamaClient.ListLocalModels();
                if (models != null && models.Any())
                {
                    availableModels = models.Select(m => m.Name).ToArray();
                    modelDropdown.Items.AddRange(availableModels);
                    modelDropdown.SelectedItem = selectedModel;  // Set the default model
                }
                else
                {
                    MessageBox.Show("No models available. Please check your Ollama installation.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading models: {ex.Message}");
            }
        }

        // Update the selected model when dropdown selection changes
        private void ModelDropdown_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedModel = modelDropdown.SelectedItem?.ToString();
            Debug.WriteLine($"Selected model updated to: {selectedModel}");
        }

        // Send the message and get a response from the selected model
        private async void SendButton_Click(object sender, EventArgs e)
        {
            string userMessage = chatInput.Text.Trim();
            if (!string.IsNullOrEmpty(userMessage))
            {
                AppendMessage("User", userMessage);
                chatInput.Clear();

                sendButton.Enabled = false;
                string aiResponse = await GetAIResponse(userMessage);

                if (InvokeRequired)
                {
                    BeginInvoke(new Action(() =>
                    {
                        AppendMessage("AI", aiResponse);
                        sendButton.Enabled = true;
                    }));
                }
                else
                {
                    AppendMessage("AI", aiResponse);
                    sendButton.Enabled = true;
                }
            }
        }

        private async Task<string> GetAIResponse(string userMessage)
        {
            string aiResponse = "";
            try
            {
                var request = new GenerateRequest
                {
                    Model = selectedModel,
                    Prompt = userMessage
                };

                // Create a placeholder message for the AI response
                AppendMessage("AI", "");

                // Stream the response and append to the same line
                await foreach (var stream in ollamaClient.Generate(request))
                {
                    aiResponse += stream.Response;

                    // Update the last message instead of appending a new one
                    UpdateLastMessage("AI", aiResponse);
                }
            }
            catch (HttpRequestException httpEx)
            {
                aiResponse = $"Error getting AI response: {httpEx.Message}";
            }
            catch (Exception ex)
            {
                aiResponse = $"Unexpected error: {ex.Message}";
            }

            return aiResponse;
        }

        // Method to append a new message (as before)
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

        // New method to update the last message
        private void UpdateLastMessage(string sender, string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => UpdateLastMessage(sender, message)));
            }
            else
            {
                // Remove the last line (the previous partial message) and append the new one
                int lastMessageIndex = chatHistory.Text.LastIndexOf($"{sender}: ");
                if (lastMessageIndex >= 0)
                {
                    chatHistory.Select(lastMessageIndex, chatHistory.Text.Length - lastMessageIndex);
                    chatHistory.SelectedText = $"{sender}: {message}\n\n";
                }
                chatHistory.ScrollToCaret();
            }
        }
    }
}
