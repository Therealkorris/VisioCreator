using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Diagnostics;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

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
        private ListView commandStatusListView;
        private Button toggleStatusButton;
        private Panel statusPanel;

        private readonly LibraryManager libraryManager;
        private readonly VisioChatManager chatManager;
        private readonly VisioCommandProcessor commandProcessor;

        public AIChatPane(string model, string apiEndpoint, string[] models, LibraryManager libraryManager)
        {
            this.libraryManager = libraryManager;

            // Initialize the chat manager, passing "this" to allow access to UpdateCommandStatus
            this.chatManager = new VisioChatManager(model, apiEndpoint, models, libraryManager, AppendToChatHistory, this);
            this.commandProcessor = new VisioCommandProcessor(Globals.ThisAddIn.Application, libraryManager);

            InitializeCustomComponents();
            PopulateModelDropdown(models);
            modelDropdown.SelectedItem = model;
        }

        private void InitializeCustomComponents()
        {
            // Chat history RichTextBox
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

            // Chat input TextBox
            chatInput = new TextBox
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                Multiline = true,
                Font = new Font("Segoe UI", 10),
            };
            chatInput.KeyDown += ChatInput_KeyDown;

            // Send button
            sendButton = new Button
            {
                Text = "Send",
                Dock = DockStyle.Bottom,
                Height = 40,
            };
            sendButton.Click += SendButton_Click;

            // Upload image button
            uploadImageButton = new Button
            {
                Text = "Upload Image",
                Dock = DockStyle.Bottom,
                Height = 40,
            };
            uploadImageButton.Click += UploadImageButton_Click;

            // Model label and dropdown
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

            // Command status ListView
            commandStatusListView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
            };
            commandStatusListView.Columns.Add("Command", 200);
            commandStatusListView.Columns.Add("Status", 100);

            // Toggle status button
            toggleStatusButton = new Button
            {
                Text = "Show/Hide Status",
                Dock = DockStyle.Bottom,
                Height = 40,
            };
            toggleStatusButton.Click += ToggleStatusButton_Click;

            // Status panel
            statusPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 300,
                Visible = false,
            };
            statusPanel.Controls.Add(commandStatusListView);

            // Add controls to the form
            Controls.Add(chatHistory);
            Controls.Add(chatInput);
            Controls.Add(uploadImageButton);
            Controls.Add(sendButton);
            Controls.Add(modelDropdown);
            Controls.Add(modelLabel);
            Controls.Add(toggleStatusButton);
            Controls.Add(statusPanel);

            // Set form properties
            Text = "AI Chat Pane";
            Width = 700;
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

        // Handles the Enter key in the chat input to send messages
        private void ChatInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !e.Shift)
            {
                e.SuppressKeyPress = true;
                sendButton.PerformClick();
            }
        }

        // Handles sending messages
        private void SendButton_Click(object sender, EventArgs e)
        {
            string userMessage = chatInput.Text.Trim();
            if (string.IsNullOrEmpty(userMessage)) return;

            chatInput.Clear();

            // Send message via VisioChatManager
            chatManager.SendMessage(userMessage);

            // Try processing the message as a JSON command
            if (IsValidJson(userMessage))
            {
                // Process the message as a JSON command
                commandProcessor.ProcessCommand(userMessage);
            }
            else
            {
                // If it's not valid JSON, still append it as plain text
                AppendToChatHistory($"You: {userMessage}");
            }

            // Update command status
            UpdateCommandStatus(userMessage, "Sent");
        }

        // Continue with the same JSON validation method
        private bool IsValidJson(string input)
        {
            input = input.Trim();
            if ((input.StartsWith("{") && input.EndsWith("}")) || (input.StartsWith("[") && input.EndsWith("]")))
            {
                try
                {
                    JToken.Parse(input);
                    return true;
                }
                catch (JsonReaderException)
                {
                    return false;
                }
            }
            return false;
        }


        // Handles uploading and sending images
        private void UploadImageButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string imagePath = openFileDialog.FileName;
                    //AppendToChatHistory("You sent an image.");

                    // Send image via VisioChatManager
                    //chatManager.SendMessageWithImage(imagePath);
                }
            }
        }

        // Handles drag-and-drop image uploads
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
                        //chatManager.SendMessageWithImage(filePath);
                    }
                }
            }
        }

        private void ChatHistory_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        // Append text to chat history
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

        // Update command status
        public void UpdateCommandStatus(string command, string status)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string, string>(UpdateCommandStatus), command, status);
            }
            else
            {
                var item = new ListViewItem(new[] { command, status });
                item.ForeColor = status == "Success" ? Color.Green : Color.Red;
                commandStatusListView.Items.Add(item);
            }
        }


        // Toggle the visibility of the status panel
        private void ToggleStatusButton_Click(object sender, EventArgs e)
        {
            statusPanel.Visible = !statusPanel.Visible;
        }

        // Placeholder for processing AI responses (e.g., adding shapes in Visio)
        private void ProcessAIResponse(string response)
        {
            // Example: Parse and execute commands from the AI response.
        }
    }
}
