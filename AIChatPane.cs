using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using OllamaSharp;
using OllamaSharp.Models;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
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
        private OllamaApiClient ollamaClient;
        private string selectedModel;
        private string[] availableModels;
        private LibraryManager libraryManager;

        public AIChatPane(Visio.Application visioApp, OllamaApiClient ollamaClient, string initialModel)
        {
            visioApplication = visioApp;
            this.ollamaClient = ollamaClient;
            this.selectedModel = initialModel;
            libraryManager = new LibraryManager(visioApp);
            libraryManager.LoadLibraries();
            InitializeComponent();
            LoadAvailableModels();
        }

        private void InitializeComponent()
        {
            // Initialize chat history
            chatHistory = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.LightYellow, // Set background color
                Font = new Font("Arial", 10),  // Set font style
                BorderStyle = BorderStyle.None
            };

            // Initialize chat input
            chatInput = new TextBox
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                Multiline = true,
                Font = new Font("Arial", 10),
                BorderStyle = BorderStyle.FixedSingle
            };
            chatInput.KeyDown += ChatInput_KeyDown; // Add key press event

            // Initialize send button
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

            // Initialize model label
            modelLabel = new Label
            {
                Text = "Select AI Model:",
                Dock = DockStyle.Top,
                Height = 20,
                Font = new Font("Arial", 10, FontStyle.Bold),
                ForeColor = Color.SteelBlue
            };

            // Initialize model dropdown
            modelDropdown = new ComboBox
            {
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Font("Arial", 10),
                DropDownStyle = ComboBoxStyle.DropDownList // Set dropdown style
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

        // Handle key press for sending message on Enter and allowing new lines on Shift+Enter
        private void ChatInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !e.Shift)
            {
                e.SuppressKeyPress = true;  // Prevent the "ding" sound
                SendMessage(); // Send the message on Enter
            }
        }

        // Send the message and get a response from the selected model
        private void SendButton_Click(object sender, EventArgs e)
        {
            SendMessage();
        }

        // Encapsulate message sending logic
        private void SendMessage()
        {
            string userMessage = chatInput.Text.Trim();
            if (!string.IsNullOrEmpty(userMessage))
            {
                AppendMessage("User", userMessage);
                chatInput.Clear();

                sendButton.Enabled = false;
                _ = GetAIResponse(userMessage);
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

                AppendMessage("AI", "");  // Create a placeholder for streaming

                await foreach (var stream in ollamaClient.Generate(request))
                {
                    aiResponse += stream.Response;

                    // Update the last message instead of appending a new one
                    UpdateLastMessage("AI", aiResponse);
                }

                // Handle AI command
                HandleAICommand(aiResponse);
            }
            catch (HttpRequestException httpEx)
            {
                aiResponse = $"Error getting AI response: {httpEx.Message}";
            }
            catch (Exception ex)
            {
                aiResponse = $"Unexpected error: {ex.Message}";
            }

            sendButton.Enabled = true;
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

        // New method to update the last message for streaming
        private void UpdateLastMessage(string sender, string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => UpdateLastMessage(sender, message)));
            }
            else
            {
                int lastMessageIndex = chatHistory.Text.LastIndexOf($"{sender}: ");
                if (lastMessageIndex >= 0)
                {
                    chatHistory.Select(lastMessageIndex, chatHistory.Text.Length - lastMessageIndex);
                    chatHistory.SelectedText = $"{sender}: {message}\n\n";
                }
                chatHistory.ScrollToCaret();
            }
        }

        // New method to handle AI commands
        private void HandleAICommand(string command)
        {
            command = command.ToLower();

            // Handle commands like "create circle", "create arrow", etc.
            if (command.Contains("create"))
            {
                if (command.Contains("circle"))
                {
                    CreateShapeFromLibrary("circle");
                }
                else if (command.Contains("arrow"))
                {
                    CreateShapeFromLibrary("arrow");
                }
                else
                {
                    AppendMessage("System", "Shape type not recognized.");
                }
            }
        }

        private void CreateShapeFromLibrary(string shapeType)
        {
            var categories = libraryManager.GetCategories();
            foreach (var category in categories)
            {
                var shapes = libraryManager.GetShapesInCategory(category);

                // Search for a shape matching the requested type (circle, arrow, etc.)
                var shape = shapes.FirstOrDefault(s => s.ToLower().Contains(shapeType));
                if (shape != null)
                {
                    libraryManager.AddShapeToDocument(category, shape, 5.0, 5.0);
                    AppendMessage("System", $"Created shape '{shape}' from category '{category}'");
                    return;
                }
            }

            AppendMessage("System", $"No shapes found matching '{shapeType}'.");
        }
    }
}
