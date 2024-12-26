using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Drawing = System.Drawing;

namespace VisioPlugin
{
    // Using fully qualified name for ShapeInfo
    public class CommandDetails
    {
        public string Id { get; set; }
        public DateTime Timestamp { get; set; }
        public string Command { get; set; }
        public string UserMessage { get; set; }
        public string AIResponse { get; set; }
        public string VisioCommand { get; set; }
        public string Status { get; set; }
        public List<VisioPlugin.ShapeInfo> AffectedShapes { get; set; }

        public CommandDetails()
        {
            Id = Guid.NewGuid().ToString("N");
            Timestamp = DateTime.Now;
            Command = "";
            UserMessage = "";
            AIResponse = "";
            VisioCommand = "";
            Status = "";
            AffectedShapes = new List<VisioPlugin.ShapeInfo>();
        }

        public override string ToString()
        {
            return $"{Command} at {Timestamp:HH:mm:ss}";
        }
    }

    public partial class AIChatPane : Form
    {
        private TextBox chatInput;
        private Button sendButton;
        private Button uploadImageButton;
        private Button clearButton;
        private RichTextBox chatHistory;
        private ComboBox modelDropdown;
        private Label modelLabel;
        private ListView commandStatusListView;
        private Button toggleStatusButton;
        private SplitContainer mainSplitContainer;
        private SplitContainer horizontalSplit;

        private readonly LibraryManager libraryManager;
        private readonly VisioChatManager chatManager;
        private readonly VisioCommandProcessor commandProcessor;
        private string pendingImagePath = null;

        // Store command details with their IDs
        private Dictionary<string, CommandDetails> commandHistory = new Dictionary<string, CommandDetails>();
        // Track the current command being processed
        private CommandDetails currentCommand = null;

        public AIChatPane(string model, string apiEndpoint, string[] models, LibraryManager libraryManager)
        {
            // Set initial form size first
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Drawing.Size(900, 600);
            this.MinimumSize = new Drawing.Size(800, 500);

            this.libraryManager = libraryManager;
            this.chatManager = new VisioChatManager(model, apiEndpoint, models, libraryManager, AppendToChatHistory, this);
            this.commandProcessor = new VisioCommandProcessor(Globals.ThisAddIn.Application, libraryManager);

            InitializeCustomComponents();
            PopulateModelDropdown(models);
            modelDropdown.SelectedItem = model;
            InitializeSplitterHandling();

            // Handle form shown event instead of load
            this.Shown += (sender, e) =>
            {
                Application.DoEvents(); // Let the form finish layout
                try
                {
                    if (horizontalSplit != null)
                    {
                        horizontalSplit.Panel1MinSize = 400;
                        horizontalSplit.Panel2MinSize = 200;
                        horizontalSplit.SplitterDistance = this.ClientSize.Width - 300;
                    }
                    if (mainSplitContainer != null)
                    {
                        mainSplitContainer.Panel1MinSize = 50;
                        mainSplitContainer.Panel2MinSize = 200;
                        mainSplitContainer.SplitterDistance = 70;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error setting splitter distances: {ex.Message}");
                }
            };
        }

        private void InitializeCustomComponents()
        {
            SuspendLayout();  // Suspend layout updates

            // Main horizontal split container between status and main area
            horizontalSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 5,  // Make splitter more visible
            };

            // Vertical split container for model selection and chat area
            mainSplitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterWidth = 5,  // Make splitter more visible
            };

            // Model selection panel (top)
            Panel modelPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5),
            };

            // Chat area panel (bottom)
            Panel chatPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5),
            };

            // Chat history RichTextBox
            chatHistory = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Drawing.Color.WhiteSmoke,
                Font = new Drawing.Font("Segoe UI", 10),
                AllowDrop = true,
                BorderStyle = BorderStyle.None,
                Margin = new Padding(5),
                SelectionIndent = 0,
            };
            chatHistory.DragDrop += ChatHistory_DragDrop;
            chatHistory.DragEnter += ChatHistory_DragEnter;

            // Bottom panel for input and buttons
            Panel bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 150,
                Padding = new Padding(5),
            };

            // Chat input TextBox
            chatInput = new TextBox
            {
                Dock = DockStyle.Top,
                Height = 100,
                Multiline = true,
                Font = new Drawing.Font("Segoe UI", 10),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(5),
            };
            chatInput.KeyDown += ChatInput_KeyDown;

            // FlowLayoutPanel for buttons
            FlowLayoutPanel buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(5),
                Margin = new Padding(5),
            };

            // Initialize buttons
            clearButton = new Button
            {
                Text = "Clear",
                Width = 100,
                Height = 40,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(5),
            };
            clearButton.Click += ClearButton_Click;

            uploadImageButton = new Button
            {
                Text = "Upload",
                Width = 100,
                Height = 40,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(5),
            };
            uploadImageButton.Click += UploadImageButton_Click;

            sendButton = new Button
            {
                Text = "Send",
                Width = 100,
                Height = 40,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(5),
            };
            sendButton.Click += SendButton_Click;

            toggleStatusButton = new Button
            {
                Text = "Status",
                Width = 100,
                Height = 40,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(5),
            };
            toggleStatusButton.Click += ToggleStatusButton_Click;

            // Model selection controls
            modelLabel = new Label
            {
                Text = "Select AI Model:",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new Drawing.Font("Segoe UI", 11, Drawing.FontStyle.Bold),
                ForeColor = Drawing.Color.SteelBlue,
                TextAlign = Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(5),
            };

            modelDropdown = new ComboBox
            {
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Drawing.Font("Segoe UI", 10),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(5),
            };
            modelDropdown.SelectedIndexChanged += ModelDropdown_SelectedIndexChanged;

            // Status panel (right side)
            commandStatusListView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Drawing.Font("Segoe UI", 10),
                OwnerDraw = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            commandStatusListView.DoubleClick += CommandStatusListView_DoubleClick;

            // Set up status list columns
            commandStatusListView.Columns.Clear();
            commandStatusListView.Columns.Add("Command", (int)(commandStatusListView.Width * 0.7));
            commandStatusListView.Columns.Add("Status", (int)(commandStatusListView.Width * 0.3));
            commandStatusListView.AllowColumnReorder = false;
            commandStatusListView.Resize += (sender, e) => ResizeListViewColumns();
            commandStatusListView.ColumnWidthChanged += (sender, e) => AdjustOtherColumnWidth(e.ColumnIndex);

            // Keep the custom drawing code for status colors
            commandStatusListView.DrawColumnHeader += (sender, e) => e.DrawDefault = true;
            commandStatusListView.DrawSubItem += (sender, e) =>
            {
                if (e.ColumnIndex == 0)
                {
                    e.DrawDefault = true;
                }
                else if (e.ColumnIndex == 1)
                {
                    Drawing.Color backgroundColor;
                    Drawing.Color borderColor;
                    
                    switch (e.Item.SubItems[1].Text.ToLower())
                    {
                        case "success":
                            backgroundColor = Drawing.Color.FromArgb(200, 255, 200);
                            borderColor = Drawing.Color.FromArgb(0, 160, 0);
                            break;
                        case "failed":
                            backgroundColor = Drawing.Color.FromArgb(255, 200, 200);
                            borderColor = Drawing.Color.FromArgb(160, 0, 0);
                            break;
                        default:
                            backgroundColor = Drawing.Color.FromArgb(255, 255, 200);
                            borderColor = Drawing.Color.FromArgb(160, 160, 0);
                            break;
                    }

                    using (var brush = new Drawing.SolidBrush(backgroundColor))
                    using (var pen = new Drawing.Pen(borderColor))
                    {
                        e.Graphics.FillRectangle(brush, e.Bounds);
                        e.Graphics.DrawRectangle(pen, e.Bounds);
                        var textBounds = new Drawing.Rectangle(e.Bounds.X + 1, e.Bounds.Y + 1, e.Bounds.Width - 2, e.Bounds.Height - 2);
                        TextRenderer.DrawText(e.Graphics, e.Item.SubItems[1].Text, e.Item.Font, textBounds, Drawing.Color.Black, 
                            TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter);
                    }
                }
            };

            // Add buttons to button panel
            buttonPanel.Controls.AddRange(new Control[] { clearButton, uploadImageButton, sendButton, toggleStatusButton });

            // Add controls to bottom panel
            bottomPanel.Controls.Add(chatInput);
            bottomPanel.Controls.Add(buttonPanel);

            // Add controls to chat panel
            chatPanel.Controls.Add(chatHistory);
            chatPanel.Controls.Add(bottomPanel);

            // Add controls to model panel
            modelPanel.Controls.Add(modelDropdown);
            modelPanel.Controls.Add(modelLabel);

            // Add panels to split containers
            mainSplitContainer.Panel1.Controls.Add(modelPanel);
            mainSplitContainer.Panel2.Controls.Add(chatPanel);
            
            horizontalSplit.Panel1.Controls.Add(mainSplitContainer);
            horizontalSplit.Panel2.Controls.Add(commandStatusListView);

            // Add the main split container to the form
            Controls.Add(horizontalSplit);

            // Set form properties
            Text = "AI Chat Pane";

            ResumeLayout(true);  // Resume layout and perform layout
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
            if (string.IsNullOrEmpty(userMessage)) return;

            // Clear the chat input first
            chatInput.Clear();

            // Only create a new command if there isn't one already being processed
            if (currentCommand == null)
            {
                currentCommand = new CommandDetails
                {
                    UserMessage = userMessage,
                    Command = "Summary of created shapes",
                    Status = "Processing"
                };

                Debug.WriteLine($"[SendButton_Click] Created new command with ID: {currentCommand.Id}");
                Debug.WriteLine($"[SendButton_Click] User Message: {userMessage}");

                // Add initial status entry
                UpdateCommandStatus(currentCommand);
            }

            // Show the user's message in chat history
            AppendToChatHistory($"You: {userMessage}");

            try
            {
                if (!string.IsNullOrEmpty(pendingImagePath))
                {
                    // Switch to the vision model for image processing
                    chatManager.SelectedModel = "llama3.2-vision:latest";
                    modelDropdown.SelectedItem = "llama3.2-vision:latest";

                    // Show the image in chat history
                    AppendImageToChatHistory(pendingImagePath);

                    try
                    {
                        // Send to image-agent with both image and message
                        await chatManager.SendImageToN8n(pendingImagePath, userMessage);
                    }
                    finally
                    {
                        pendingImagePath = null; // Reset pending image
                    }
                }
                else
                {
                    // Regular text message - send to chat-agent
                    await chatManager.SendMessage(userMessage);
                }
            }
            catch (Exception ex)
            {
                AppendToChatHistory($"Error: {ex.Message}");
                Debug.WriteLine($"[Error] Sending message: {ex.Message}");

                if (currentCommand != null)
                {
                    currentCommand.Status = "Failed";
                    currentCommand.AIResponse = $"Error: {ex.Message}";
                    UpdateCommandStatus(currentCommand);
                    currentCommand = null;
                }
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
                    // Store the image path for later sending
                    pendingImagePath = imagePath;
                    
                    // Just update the chat input to show the image is ready
                    chatInput.Text = $"[📎 {Path.GetFileName(imagePath)}] {chatInput.Text}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error adding image: {ex.Message}");
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

        // Append text to chat history with styling
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
                    int start = chatHistory.TextLength;
                    chatHistory.SelectionStart = start;
                    chatHistory.SelectionIndent = 0;
                    
                    // Set colors and style based on message type
                    if (message.StartsWith("You: "))
                    {
                        chatHistory.SelectionColor = Drawing.Color.RoyalBlue;
                        chatHistory.SelectionFont = new Drawing.Font(chatHistory.Font, Drawing.FontStyle.Bold);
                    }
                    else if (message.StartsWith("AI: "))
                    {
                        // Store AI response in current command if available
                        if (currentCommand != null)
                        {
                            currentCommand.AIResponse = message.Substring(4); // Remove "AI: " prefix
                        }

                        chatHistory.SelectionColor = Drawing.Color.FromArgb(128, 0, 128); // Purple
                        chatHistory.SelectionFont = new Drawing.Font(chatHistory.Font, Drawing.FontStyle.Regular);
                    }
                    else
                    {
                        chatHistory.SelectionColor = Drawing.Color.Gray;
                        chatHistory.SelectionFont = new Drawing.Font(chatHistory.Font, Drawing.FontStyle.Italic);
                    }

                    chatHistory.AppendText(message + Environment.NewLine);
                    chatHistory.ScrollToCaret();
                }
            }
        }

        // Append image to chat history with preview
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
                    // Add a separator line
                    chatHistory.AppendText("----------------------------------------" + Environment.NewLine);
                    
                    // Add image path with styling
                    int pathStart = chatHistory.TextLength;
                    chatHistory.SelectionStart = pathStart;
                    chatHistory.SelectionColor = Drawing.Color.Gray;
                    chatHistory.SelectionFont = new Drawing.Font(chatHistory.Font, Drawing.FontStyle.Italic);
                    chatHistory.AppendText($"Image: {Path.GetFileName(imagePath)}" + Environment.NewLine);

                    using (Drawing.Image image = Drawing.Image.FromFile(imagePath))
                    {
                        float aspectRatio = (float)image.Width / image.Height;
                        int maxWidth = chatHistory.ClientSize.Width - 40;
                        int maxHeight = 200;

                        int newWidth = Math.Min(image.Width, maxWidth);
                        int newHeight = (int)(newWidth / aspectRatio);

                        if (newHeight > maxHeight)
                        {
                            newHeight = maxHeight;
                            newWidth = (int)(newHeight * aspectRatio);
                        }

                        using (Drawing.Image resizedImage = new Drawing.Bitmap(image, new Drawing.Size(newWidth, newHeight)))
                        {
                            chatHistory.AppendText(Environment.NewLine);
                            Clipboard.SetImage(resizedImage);
                            chatHistory.ReadOnly = false;
                            chatHistory.SelectionStart = chatHistory.TextLength;
                            chatHistory.Paste();
                            chatHistory.ReadOnly = true;
                            chatHistory.AppendText(Environment.NewLine + Environment.NewLine);
                        }
                    }

                    // Add another separator line
                    chatHistory.AppendText("----------------------------------------" + Environment.NewLine);
                    chatHistory.ScrollToCaret();
                }
                catch (Exception ex)
                {
                    AppendToChatHistory($"Error loading image: {ex.Message}");
                    Debug.WriteLine($"[Error] Loading image: {ex.Message}");
                }
            }
        }

        // Update command status with more details
        public void UpdateCommandStatus(CommandDetails details)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<CommandDetails>(UpdateCommandStatus), details);
                return;
            }

            Debug.WriteLine($"[UpdateCommandStatus] Updating command: {details.Id}");
            Debug.WriteLine($"[UpdateCommandStatus] Status: {details.Status}");
            Debug.WriteLine($"[UpdateCommandStatus] Affected shapes count: {details.AffectedShapes?.Count ?? 0}");

            // Create a deep copy of the command details
            var detailsCopy = new CommandDetails
            {
                Id = details.Id,
                Timestamp = details.Timestamp,
                Command = details.Command,
                UserMessage = details.UserMessage,
                AIResponse = details.AIResponse,
                VisioCommand = details.VisioCommand,
                Status = details.Status
            };

            // Make a deep copy of affected shapes
            if (details.AffectedShapes != null)
            {
                detailsCopy.AffectedShapes = details.AffectedShapes.Select(s => new ShapeInfo
                {
                    ShapeId = s.ShapeId,
                    ShapeType = s.ShapeType,
                    ShapeColor = s.ShapeColor
                }).ToList();
                Debug.WriteLine($"[UpdateCommandStatus] Copied {detailsCopy.AffectedShapes.Count} shapes to command history");
            }

            // Store command in history
            commandHistory[detailsCopy.Id] = detailsCopy;

            // Update list view
            ListViewItem existingItem = null;
            foreach (ListViewItem item in commandStatusListView.Items)
            {
                if (item.Tag?.ToString() == detailsCopy.Id)
                {
                    existingItem = item;
                    break;
                }
            }

            string displayText = detailsCopy.Command;

            if (existingItem != null)
            {
                existingItem.SubItems[0].Text = displayText;
                existingItem.SubItems[1].Text = detailsCopy.Status;
            }
            else
            {
                var item = new ListViewItem(new[] { displayText, detailsCopy.Status }) { Tag = detailsCopy.Id };
                commandStatusListView.Items.Insert(0, item);
                item.EnsureVisible();
            }

            commandStatusListView.Refresh();
        }

        // Toggle the visibility of the status panel
        private void ToggleStatusButton_Click(object sender, EventArgs e)
        {
            if (horizontalSplit != null)
            {
                bool isStatusVisible = horizontalSplit.Panel2Collapsed;
                horizontalSplit.Panel2Collapsed = !isStatusVisible;
                toggleStatusButton.Text = isStatusVisible ? "Hide Status" : "Show Status";
            }
        }

        // Store the last splitter distance when hiding the panel
        private int lastSplitterDistance = 0;

        // Handle splitter movement
        private void InitializeSplitterHandling()
        {
            if (horizontalSplit != null)
            {
                lastSplitterDistance = horizontalSplit.SplitterDistance;
                
                horizontalSplit.SplitterMoved += (sender, e) =>
                {
                    if (!horizontalSplit.Panel2Collapsed)
                    {
                        lastSplitterDistance = horizontalSplit.SplitterDistance;
                    }
                };
            }
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

        // Clear button click handler
        private void ClearButton_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to clear all chat history and status?", "Confirm Clear",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                chatHistory.Clear();
                chatInput.Clear();
                commandStatusListView.Items.Clear();
                pendingImagePath = null;
            }
        }

        // Improved double-click handler for status items
        private void CommandStatusListView_DoubleClick(object sender, EventArgs e)
        {
            if (commandStatusListView.SelectedItems.Count == 0) return;

            var item = commandStatusListView.SelectedItems[0];
            string commandId = item.Tag?.ToString();
            if (string.IsNullOrEmpty(commandId) || !commandHistory.ContainsKey(commandId)) return;

            ShowCommandDetails(commandId);
        }

        private void ShowCommandDetails(string commandId)
        {
            if (!commandHistory.TryGetValue(commandId, out CommandDetails details))
            {
                Debug.WriteLine($"[ShowCommandDetails] Could not find command details for ID: {commandId}");
                return;
            }

            Debug.WriteLine($"[ShowCommandDetails] Showing details for command: {commandId}");
            Debug.WriteLine($"[ShowCommandDetails] Affected shapes count: {details.AffectedShapes?.Count ?? 0}");
            if (details.AffectedShapes?.Any() == true)
            {
                foreach (var shape in details.AffectedShapes)
                {
                    Debug.WriteLine($"[ShowCommandDetails] Shape to display: {shape}");
                }
            }

            var dialog = new Form
            {
                Text = $"Command Details - {details.Command}",
                Size = new Drawing.Size(600, 500),
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false,
                MaximizeBox = false,
                FormBorderStyle = FormBorderStyle.FixedDialog
            };

            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 6,
                Padding = new Padding(10),
                CellBorderStyle = TableLayoutPanelCellBorderStyle.None
            };

            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));

            // Add header labels with consistent styling
            var timestampLabel = new Label { Text = "Timestamp:", Font = new Drawing.Font(Font.FontFamily, 9, Drawing.FontStyle.Bold), Dock = DockStyle.Fill };
            var timestampValue = new Label { Text = details.Timestamp.ToString("yyyy-MM-dd HH:mm:ss"), Font = new Drawing.Font(Font.FontFamily, 9), Dock = DockStyle.Fill };
            
            var statusLabel = new Label { Text = "Status:", Font = new Drawing.Font(Font.FontFamily, 9, Drawing.FontStyle.Bold), Dock = DockStyle.Fill };
            var statusValue = new Label { Text = details.Status, Font = new Drawing.Font(Font.FontFamily, 9), Dock = DockStyle.Fill };
            statusValue.ForeColor = details.Status.ToLower() == "success" ? Drawing.Color.Green : Drawing.Color.Red;

            // Add message sections
            var userMessageLabel = new Label { Text = "User Message:", Font = new Drawing.Font(Font.FontFamily, 9, Drawing.FontStyle.Bold), Dock = DockStyle.Fill };
            var userMessageBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Height = 60,
                Dock = DockStyle.Fill,
                Text = details.UserMessage ?? ""
            };

            var aiResponseLabel = new Label { Text = "AI Response:", Font = new Drawing.Font(Font.FontFamily, 9, Drawing.FontStyle.Bold), Dock = DockStyle.Fill };
            var aiResponseBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Height = 60,
                Dock = DockStyle.Fill,
                Text = details.AIResponse ?? ""
            };

            // Add affected shapes section
            var affectedShapesLabel = new Label { Text = "Affected Shapes:", Font = new Drawing.Font(Font.FontFamily, 9, Drawing.FontStyle.Bold), Dock = DockStyle.Fill };
            var affectedShapesBox = new ListBox
            {
                Height = 100,
                Dock = DockStyle.Fill
            };

            if (details.AffectedShapes?.Any() == true)
            {
                Debug.WriteLine($"[ShowCommandDetails] Adding {details.AffectedShapes.Count} shapes to ListBox");
                foreach (var shape in details.AffectedShapes)
                {
                    Debug.WriteLine($"[ShowCommandDetails] Adding to ListBox: {shape}");
                    affectedShapesBox.Items.Add(shape.ToString());
                }
            }
            else
            {
                Debug.WriteLine("[ShowCommandDetails] No shapes to display, adding 'No shapes affected' message");
                affectedShapesBox.Items.Add("No shapes affected");
            }

            // Set up the layout
            mainPanel.Controls.Add(timestampLabel, 0, 0);
            mainPanel.Controls.Add(timestampValue, 1, 0);
            mainPanel.Controls.Add(statusLabel, 0, 1);
            mainPanel.Controls.Add(statusValue, 1, 1);
            
            mainPanel.Controls.Add(userMessageLabel, 0, 2);
            mainPanel.SetColumnSpan(userMessageBox, 2);
            mainPanel.Controls.Add(userMessageBox, 0, 3);
            
            mainPanel.Controls.Add(aiResponseLabel, 0, 4);
            mainPanel.SetColumnSpan(aiResponseBox, 2);
            mainPanel.Controls.Add(aiResponseBox, 0, 5);

            mainPanel.Controls.Add(affectedShapesLabel, 0, 6);
            mainPanel.SetColumnSpan(affectedShapesBox, 2);
            mainPanel.Controls.Add(affectedShapesBox, 0, 7);

            dialog.Controls.Add(mainPanel);
            dialog.ShowDialog();
        }

        public CommandDetails GetCurrentCommand()
        {
            Debug.WriteLine($"[GetCurrentCommand] Current command: {currentCommand?.Id ?? "null"}");
            return currentCommand;
        }

        public void ResetCurrentCommand()
        {
            Debug.WriteLine($"[ResetCurrentCommand] Resetting command: {currentCommand?.Id ?? "null"}");
            currentCommand = null;
        }

        private string ExtractChatMessage(JObject responseObject)
        {
            try
            {
                // Get the first property which contains the main message
                var firstProperty = responseObject.Properties().FirstOrDefault();
                if (firstProperty == null) return "";

                // If the value is a simple string, return it
                if (firstProperty.Value.Type == JTokenType.String)
                    return firstProperty.Value.ToString();

                // If it's a nested structure, try to find the deepest string value
                var message = firstProperty.Name;
                var currentToken = firstProperty.Value;

                // Try to find the actual message in the nested structure
                while (currentToken is JObject obj)
                {
                    var emptyKeyProp = obj.Properties().FirstOrDefault(p => p.Name == "");
                    if (emptyKeyProp != null && emptyKeyProp.Value.Type == JTokenType.String)
                    {
                        return emptyKeyProp.Value.ToString();
                    }

                    var firstProp = obj.Properties().FirstOrDefault();
                    if (firstProp == null) break;

                    if (!string.IsNullOrWhiteSpace(firstProp.Name) && firstProp.Name != " ")
                    {
                        message = firstProp.Name;
                    }

                    currentToken = firstProp.Value;
                }

                // Clean up the message
                message = message.Replace("\n", " ").Trim();
                Debug.WriteLine($"[ExtractChatMessage] Extracted message: {message}");
                return message;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExtractChatMessage] Error extracting message: {ex.Message}");
                return "";
            }
        }

    }
}