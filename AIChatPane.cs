using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Diagnostics;

namespace VisioPlugin
{
    public partial class AIChatPane : Form
    {
        private TextBox chatInput;
        private Button sendButton;
        private Button uploadImageButton;
        private RichTextBox chatHistory;
        private ComboBox modelDropdown;
        private Label modelLabel;

        private readonly LibraryManager libraryManager;
        private readonly string apiEndpoint;
        private string selectedModel;

        public AIChatPane(string model, string apiEndpoint, string[] models, LibraryManager libraryManager)
        {
            this.libraryManager = libraryManager;
            this.apiEndpoint = apiEndpoint;
            this.selectedModel = model;

            InitializeCustomComponents();

            // Populate models into the dropdown
            PopulateModelDropdown(models);
            modelDropdown.SelectedItem = model;
        }

        private void InitializeCustomComponents()
        {
            // Initialize controls
            chatHistory = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.WhiteSmoke,
                Font = new Font("Segoe UI", 10),
                AllowDrop = true,
            };
            chatHistory.DragDrop += ChatHistory_DragDrop;
            chatHistory.DragEnter += ChatHistory_DragEnter;

            chatInput = new TextBox
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                Multiline = true,
                Font = new Font("Segoe UI", 10),
            };
            chatInput.KeyDown += ChatInput_KeyDown;

            sendButton = new Button
            {
                Text = "Send",
                Dock = DockStyle.Bottom,
                Height = 40,
            };
            sendButton.Click += SendButton_Click;

            uploadImageButton = new Button
            {
                Text = "Upload Image",
                Dock = DockStyle.Bottom,
                Height = 40,
            };
            uploadImageButton.Click += UploadImageButton_Click;

            modelLabel = new Label
            {
                Text = "Select AI Model:",
                Dock = DockStyle.Top,
                Height = 20,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.SteelBlue,
            };

            modelDropdown = new ComboBox
            {
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Font("Segoe UI", 10),
                DropDownStyle = ComboBoxStyle.DropDownList,
            };
            modelDropdown.SelectedIndexChanged += ModelDropdown_SelectedIndexChanged;

            // Add controls to the form
            Controls.Add(chatHistory);
            Controls.Add(chatInput);
            Controls.Add(uploadImageButton);
            Controls.Add(sendButton);
            Controls.Add(modelDropdown);
            Controls.Add(modelLabel);

            // Set form properties
            Text = "AI Chat Pane";
            Width = 400;
            Height = 600;
        }

        private void PopulateModelDropdown(string[] models)
        {
            modelDropdown.Items.Clear();

            if (models != null && models.Length > 0)
            {
                modelDropdown.Items.AddRange(models);
            }
            else
            {
                MessageBox.Show("No models available.");
            }
        }

        private void ChatInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !e.Shift)
            {
                e.SuppressKeyPress = true;
                sendButton.PerformClick();
            }
        }

        private async void SendButton_Click(object sender, EventArgs e)
        {
            string userMessage = chatInput.Text.Trim();
            if (string.IsNullOrEmpty(userMessage)) return;

            chatInput.Clear();
            AppendToChatHistory($"You: {userMessage}");

            try
            {
                var response = await BackendCommunication.SendMessage(apiEndpoint, userMessage, selectedModel);
                AppendToChatHistory($"AI: {response}");

                // Process AI commands if any
                ProcessAIResponse(response);
            }
            catch (Exception ex)
            {
                AppendToChatHistory($"Error: {ex.Message}");
            }
        }

        private async void UploadImageButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string imagePath = openFileDialog.FileName;
                    AppendToChatHistory("You sent an image.");

                    try
                    {
                        var response = await BackendCommunication.SendImage(apiEndpoint, imagePath, selectedModel);
                        AppendToChatHistory($"AI: {response}");

                        // Process AI commands if any
                        ProcessAIResponse(response);
                    }
                    catch (Exception ex)
                    {
                        AppendToChatHistory($"Error: {ex.Message}");
                    }
                }
            }
        }

        private void ChatHistory_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string filePath = files[0];
                    if (filePath.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) || filePath.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
                    {
                        UploadImage(filePath);
                    }
                }
            }
        }

        private void ChatHistory_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private async void UploadImage(string imagePath)
        {
            AppendToChatHistory("You sent an image.");

            try
            {
                var response = await BackendCommunication.SendImage(apiEndpoint, imagePath, selectedModel);
                AppendToChatHistory($"AI: {response}");

                // Process AI commands if any
                ProcessAIResponse(response);
            }
            catch (Exception ex)
            {
                AppendToChatHistory($"Error: {ex.Message}");
            }
        }

        private void AppendToChatHistory(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(AppendToChatHistory), message);
            }
            else
            {
                chatHistory.AppendText(message + Environment.NewLine);
                chatHistory.ScrollToCaret();
            }
        }

        private void ModelDropdown_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedModel = modelDropdown.SelectedItem?.ToString();
            AppendToChatHistory($"Model changed to: {selectedModel}");
        }

        private void ProcessAIResponse(string response)
        {
            // If the AI response includes commands for Visio, process them here
            // For example, parse the response and use LibraryManager to add shapes
            // This is a placeholder for actual implementation

            // Example:
            /*
            var aiCommand = JsonConvert.DeserializeObject<AICommand>(response);
            if (aiCommand != null)
            {
                libraryManager.AddShapeToDocument(
                    aiCommand.Category,
                    aiCommand.ShapeName,
                    aiCommand.Position.X,
                    aiCommand.Position.Y,
                    aiCommand.Size.Width,
                    aiCommand.Size.Height
                );
            }
            */
        }
    }
}
