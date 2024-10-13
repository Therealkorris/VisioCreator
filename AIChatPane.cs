using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.IO;
using System.Net.Http;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioPlugin
{
    public partial class AIChatPane : Form
    {
        private TextBox chatInput;
        private Button sendButton, uploadImageButton;
        private RichTextBox chatHistory;
        private ComboBox modelDropdown;
        private Label modelLabel;
        private readonly HttpClient httpClient = new HttpClient();
        private string selectedModel;
        private string pythonApiEndpoint;
        private string[] availableModels;
        private LibraryManager libraryManager;

        // Constructor now accepts available models from ThisAddIn
        public AIChatPane(string model, string apiEndpoint, string[] models, LibraryManager libraryManager)
        {
            selectedModel = model;
            pythonApiEndpoint = apiEndpoint;
            availableModels = models; // Use the models passed from ThisAddIn
            this.libraryManager = libraryManager;
            InitializeComponent();
            PopulateModelDropdown();  // Populate dropdown with models
        }

        private void InitializeComponent()
        {
            chatHistory = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowDrop = true,
                BackColor = Color.LightYellow,
                Font = new Font("Arial", 10),
            };
            chatHistory.DragDrop += ChatHistory_DragDrop;
            chatHistory.DragEnter += ChatHistory_DragEnter;

            chatInput = new TextBox
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                Multiline = true,
                Font = new Font("Arial", 10),
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

            // Model selection dropdown and label
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

            // Add controls to the form
            Controls.Add(chatHistory);
            Controls.Add(chatInput);
            Controls.Add(uploadImageButton);
            Controls.Add(sendButton);
            Controls.Add(modelDropdown);
            Controls.Add(modelLabel);
        }

        // Populate the model dropdown with the available models from the ThisAddIn's list
        private void PopulateModelDropdown()
        {
            modelDropdown.Items.Clear();

            if (availableModels != null && availableModels.Length > 0)
            {
                modelDropdown.Items.AddRange(availableModels);
                modelDropdown.SelectedItem = selectedModel;
                Debug.WriteLine($"Models loaded into dropdown: {string.Join(", ", availableModels)}");
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
            await SendMessage();
        }

        private async void UploadImageButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    await SendMessageWithImage(openFileDialog.FileName);
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
                    if (filePath.EndsWith(".jpg") || filePath.EndsWith(".png"))
                    {
                        Task.Run(() => SendMessageWithImage(filePath));
                    }
                }
            }
        }

        private void ChatHistory_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        // Send a message to the AI and accumulate the response
        private async Task SendMessage()
        {
            try
            {
                string userMessage = chatInput.Text.Trim();
                if (string.IsNullOrEmpty(userMessage)) return;

                AppendToChatHistory("User: " + userMessage);
                chatInput.Clear();

                var content = new MultipartFormDataContent();
                content.Add(new StringContent(userMessage), "prompt");
                content.Add(new StringContent(selectedModel), "model");

                // Log before sending message
                Debug.WriteLine($"Sending message to backend: {userMessage}");

                var response = await httpClient.PostAsync($"{pythonApiEndpoint}/test-visio-command", content);
                var responseString = await response.Content.ReadAsStringAsync();

                // Log after receiving response
                Debug.WriteLine($"Received response from backend: {responseString}");

                AppendToChatHistory("AI: " + responseString.Trim());

                // Parse the response and execute the Visio command
                var commandResponse = JsonConvert.DeserializeObject<dynamic>(responseString);
                if (commandResponse.status == "success" && commandResponse.command != null)
                {
                    await ExecuteVisioCommand(
                        commandResponse.command.shape.ToString(),
                        (float)commandResponse.command.x,
                        (float)commandResponse.command.y,
                        (float)commandResponse.command.width,
                        (float)commandResponse.command.height
                    );
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
            AppendToChatHistory("User sent an image: " + Path.GetFileName(imagePath));

            try
            {
                var content = new MultipartFormDataContent();

                // Add the image file
                var imageContent = new ByteArrayContent(File.ReadAllBytes(imagePath));
                imageContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("image/jpeg");
                content.Add(imageContent, "file", Path.GetFileName(imagePath));

                // Add the prompt and model as form data
                content.Add(new StringContent("Image analysis prompt"), "prompt");
                content.Add(new StringContent(selectedModel), "model");

                // Send the request to FastAPI server
                var response = await httpClient.PostAsync($"{pythonApiEndpoint}/image-prompt", content);

                var responseContent = await response.Content.ReadAsStringAsync();
                AppendToChatHistory("AI: " + responseContent);
                Debug.WriteLine($"AI Response: {responseContent}");
            }
            catch (Exception ex)
            {
                AppendToChatHistory("Error: " + ex.Message);
                Debug.WriteLine("Error sending image: " + ex.Message);
            }
        }

        // Execute Visio command via BackendCommunication
        private Task ExecuteVisioCommand(string shape, float x, float y, float width, float height)
        {
            try
            {
                // Find the appropriate category for the shape
                string category = FindCategoryForShape(shape);
                if (string.IsNullOrEmpty(category))
                {
                    AppendToChatHistory($"Error: Shape '{shape}' not found in any category.");
                    return Task.CompletedTask;
                }

                // Add the shape to the document using LibraryManager
                libraryManager.AddShapeToDocument(category, shape, x, y);

                // Get the added shape (assuming it's the last shape added to the active page)
                Visio.Shape newShape = Globals.ThisAddIn.Application.ActivePage.Shapes.Cast<Visio.Shape>().Last();

                // Resize the shape
                newShape.Resize(Visio.VisResizeDirection.visResizeDirE, width / newShape.Cells["Width"].ResultIU, Visio.VisUnitCodes.visInches);
                newShape.Resize(Visio.VisResizeDirection.visResizeDirN, height / newShape.Cells["Height"].ResultIU, Visio.VisUnitCodes.visInches);

                AppendToChatHistory($"Visio Command Executed: {shape} created successfully at ({x}, {y}) with dimensions {width}x{height}");
            }
            catch (Exception ex)
            {
                AppendToChatHistory("Error executing Visio command: " + ex.Message);
                System.Diagnostics.Debug.WriteLine("Error executing Visio command: " + ex.Message);
            }

            return Task.CompletedTask;
        }

        private string FindCategoryForShape(string shapeName)
        {
            foreach (var category in libraryManager.GetCategories())
            {
                var shapes = libraryManager.GetShapesInCategory(category);
                if (shapes.Any(s => string.Equals(s, shapeName, StringComparison.OrdinalIgnoreCase)))
                {
                    return category;
                }
            }
            return null;
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
            Debug.WriteLine($"Selected model updated to: {selectedModel}");
        }

        public class ModelResponse
        {
            public List<string> Models { get; set; }
        }
    }
}
