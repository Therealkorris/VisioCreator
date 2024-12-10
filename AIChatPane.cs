using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Diagnostics;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Text; // Added for Encoding
using System.Net.Http; // Added for HttpClient
using System.IO;

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
        private string pendingImagePath = null;

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
                BorderStyle = BorderStyle.None, // Remove borders for a cleaner look
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
                Padding = new Padding(10), // Add padding for better input area look
                BorderStyle = BorderStyle.FixedSingle
            };
            chatInput.KeyDown += ChatInput_KeyDown;

            // FlowLayoutPanel to contain buttons for better layout
            FlowLayoutPanel buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Padding = new Padding(5),
                Margin = new Padding(5),
            };

            // Send button
            sendButton = new Button
            {
                Text = "Send",
                Width = 100, // Adjusted width for better text fitting
                Height = 40,
                FlatStyle = FlatStyle.Flat, // Modern flat style
                Margin = new Padding(10, 5, 5, 5), // Spacing around the button
                Padding = new Padding(5), // Padding inside the button
            };
            sendButton.Click += SendButton_Click;

            // Upload image button
            uploadImageButton = new Button
            {
                Text = "Upload", // Renamed to fit better
                Width = 100,
                Height = 40,
                FlatStyle = FlatStyle.Flat, // Modern flat style
                Margin = new Padding(5), // Spacing around the button
                Padding = new Padding(5), // Padding inside the button
            };
            uploadImageButton.Click += UploadImageButton_Click;

            // Toggle status button
            toggleStatusButton = new Button
            {
                Text = "Status", // Renamed to fit better
                Width = 130,
                Height = 40,
                FlatStyle = FlatStyle.Flat, // Modern flat style
                Margin = new Padding(5), // Spacing around the button
                Padding = new Padding(5), // Padding inside the button
            };
            toggleStatusButton.Click += ToggleStatusButton_Click;

            // Add buttons to FlowLayoutPanel
            buttonPanel.Controls.Add(uploadImageButton);
            buttonPanel.Controls.Add(sendButton);
            buttonPanel.Controls.Add(toggleStatusButton);

            // Model label and dropdown
            modelLabel = new Label
            {
                Text = "Select AI Model:",
                Dock = DockStyle.Top,
                Height = 25, // Adjusted height for better readability
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.SteelBlue,
                TextAlign = ContentAlignment.MiddleLeft, // Align text properly
                Padding = new Padding(10, 5, 0, 5) // Top-left padding for spacing
            };

            modelDropdown = new ComboBox
            {
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Font("Segoe UI", 10),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(10, 0, 10, 10), // Better margin for dropdown
            };
            modelDropdown.SelectedIndexChanged += ModelDropdown_SelectedIndexChanged; // Add event handler

            // Command status ListView
            commandStatusListView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 10),
                OwnerDraw = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };

            // Clear existing columns and add only two columns
            commandStatusListView.Columns.Clear();
            commandStatusListView.Columns.Add("Command", (int)(commandStatusListView.Width * 0.7));
            commandStatusListView.Columns.Add("Status", (int)(commandStatusListView.Width * 0.3));

            // Allow column reordering and resizing
            commandStatusListView.AllowColumnReorder = false;

            // Resize columns to fill the ListView width
            commandStatusListView.Resize += (sender, e) => ResizeListViewColumns();

            // Handle column width changes
            commandStatusListView.ColumnWidthChanged += (sender, e) => AdjustOtherColumnWidth(e.ColumnIndex);

            // Custom drawing event to handle status color and borders
            commandStatusListView.DrawColumnHeader += (sender, e) => e.DrawDefault = true;
            commandStatusListView.DrawSubItem += (sender, e) =>
            {
                if (e.ColumnIndex == 0)
                {
                    e.DrawDefault = true;
                }
                else if (e.ColumnIndex == 1)
                {
                    if (e.Item.SubItems[1].Text == "Success")
                    {
                        e.Graphics.FillRectangle(Brushes.LightGreen, e.Bounds);
                        e.Graphics.DrawRectangle(Pens.Green, e.Bounds);
                    }
                    else
                    {
                        e.Graphics.FillRectangle(Brushes.LightCoral, e.Bounds);
                        e.Graphics.DrawRectangle(Pens.Red, e.Bounds);
                    }
                    TextRenderer.DrawText(e.Graphics, e.Item.SubItems[1].Text, e.Item.Font, e.Bounds, Color.Black, TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter);
                }
            };

            // Status panel
            statusPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 300,
                Visible = true,
                Padding = new Padding(5),
                BorderStyle = BorderStyle.FixedSingle // Clean border to distinguish panel
            };
            statusPanel.Controls.Add(commandStatusListView);

            // Add controls to the form
            Controls.Add(chatHistory);
            Controls.Add(chatInput);
            Controls.Add(buttonPanel); // Add button panel back to the form
            Controls.Add(modelDropdown);
            Controls.Add(modelLabel);
            Controls.Add(statusPanel);

            // Set form properties
            Text = "AI Chat Pane";
            Width = 700;
            Height = 600;
            MinimumSize = new System.Drawing.Size(600, 500); // Enforce minimum size for better usability
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
        private async void SendButton_Click(object sender, EventArgs e)
        {
            string userMessage = chatInput.Text.Trim();

            if (!string.IsNullOrEmpty(pendingImagePath))
            {
                // If there's a pending image, prepend image information to the message
                userMessage = $"Image: {Path.GetFileName(pendingImagePath)}\n{userMessage}";

                // Switch to the vision model
                chatManager.SelectedModel = "llama3.2-vision:latest";
                modelDropdown.SelectedItem = "llama3.2-vision:latest";

                // Send the image
                try
                {
                    await chatManager.SendImageToN8n(pendingImagePath);
                }
                catch (Exception ex)
                {
                    AppendToChatHistory($"Error sending image: {ex.Message}");
                    Debug.WriteLine($"[Error] Sending image: {ex.Message}");
                }
                finally
                {
                    pendingImagePath = null; // Reset pending image
                }
            }

            // Clear the chat input
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => chatInput.Clear()));
            }
            else
            {
                chatInput.Clear();
            }

            AppendToChatHistory($"You: {userMessage}");

            try
            {
                // Send message via VisioChatManager
                await chatManager.SendMessage(userMessage);
            }
            catch (Exception ex)
            {
                AppendToChatHistory($"Error: {ex.Message}");
                Debug.WriteLine($"[Error] Sending message: {ex.Message}");
                UpdateCommandStatus(userMessage, "Failed");
            }
        }

        private void ModelDropdown_SelectedIndexChanged(object sender, EventArgs e)
        {
            chatManager.SelectedModel = modelDropdown.SelectedItem.ToString();
            AppendToChatHistory($"Model changed to: {modelDropdown.SelectedItem.ToString()}");
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

                    // Instead of appending the image to the chat history, add it to the chat input box
                    AddImageToChatInput(imagePath);

                    // Store the path for later sending
                    pendingImagePath = imagePath;
                }
            }
        }

        private void AddImageToChatInput(string imagePath)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(AddImageToChatInput), imagePath);
            }
            else
            {
                try
                {
                    // Add the image to the chat input box instead of the chat history
                    chatInput.Text = $"[Image: {Path.GetFileName(imagePath)}] {chatInput.Text}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error adding image to chat input: {ex.Message}");
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
                    if (filePath.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                        filePath.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) ||
                        filePath.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
                    {
                        // Add the image to the chat input box instead of the chat history
                        AddImageToChatInput(filePath);

                        // Store the path for later sending
                        pendingImagePath = filePath;

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
        public void AppendToChatHistory(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(AppendToChatHistory), message);
            }
            else
            {
                if (!chatHistory.IsDisposed)
                {
                    chatHistory.AppendText(message + Environment.NewLine);
                    chatHistory.ScrollToCaret();
                }
            }
        }

        // Append image to chat history
        private void AppendImageToChatHistory(string imagePath)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(AppendImageToChatHistory), imagePath);
            }
            else
            {
                try
                {
                    Image image = Image.FromFile(imagePath);
                    float aspectRatio = (float)image.Width / image.Height;
                    int maxWidth = chatHistory.ClientSize.Width - 20; // Adjust for padding
                    int maxHeight = chatHistory.ClientSize.Height / 3; // Limit height

                    // Calculate new dimensions
                    int newWidth = image.Width;
                    int newHeight = image.Height;

                    if (newWidth > maxWidth)
                    {
                        newWidth = maxWidth;
                        newHeight = (int)(newWidth / aspectRatio);
                    }

                    if (newHeight > maxHeight)
                    {
                        newHeight = maxHeight;
                        newWidth = (int)(newHeight * aspectRatio);
                    }

                    // Resize the image
                    Image resizedImage = new Bitmap(image, new Size(newWidth, newHeight));

                    // Insert the resized image into the chat history
                    Clipboard.SetImage(resizedImage);
                    chatHistory.ReadOnly = false;
                    chatHistory.SelectionStart = chatHistory.TextLength;
                    chatHistory.Paste();
                    chatHistory.ReadOnly = true;
                    chatHistory.AppendText(Environment.NewLine);
                    chatHistory.ScrollToCaret();

                    image.Dispose();
                    resizedImage.Dispose();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading image: {ex.Message}");
                }
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
                item.ForeColor = status == "Success" ? Color.Green : Color.Red; // Conditional coloring
                commandStatusListView.Items.Add(item);
            }
        }

        // Toggle the visibility of the status panel
        private void ToggleStatusButton_Click(object sender, EventArgs e)
        {
            statusPanel.Visible = !statusPanel.Visible;
            toggleStatusButton.Text = statusPanel.Visible ? "Hide Status" : "Show Status"; // Update button text
        }

        private void ResizeListViewColumns()
        {
            if (commandStatusListView.Columns.Count == 2)
            {
                int totalWidth = commandStatusListView.ClientSize.Width;
                commandStatusListView.Columns[0].Width = (int)(totalWidth * 0.7);
                commandStatusListView.Columns[1].Width = (int)(totalWidth * 0.3);
            }
        }

        private void AdjustOtherColumnWidth(int changedColumnIndex)
        {
            if (commandStatusListView.Columns.Count != 2) return;

            int totalWidth = commandStatusListView.ClientSize.Width;
            int changedColumnWidth = commandStatusListView.Columns[changedColumnIndex].Width;
            int otherColumnIndex = 1 - changedColumnIndex;

            commandStatusListView.Columns[otherColumnIndex].Width = totalWidth - changedColumnWidth;
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

    }
}