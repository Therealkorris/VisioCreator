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
using System.Net.Http.Headers;

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

        // Constructor now accepts available models from ThisAddIn
        public AIChatPane(string model, string apiEndpoint, string[] models)
        {
            selectedModel = model;
            pythonApiEndpoint = apiEndpoint;
            availableModels = models; // Use the models passed from ThisAddIn
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

        // Now we're using the available models from the ThisAddIn's list
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

                // Accumulate the response instead of processing chunks individually
                var responseStream = await response.Content.ReadAsStreamAsync();
                using (var reader = new StreamReader(responseStream))
                {
                    StringBuilder fullResponse = new StringBuilder();
                    string line;

                    // Read the response and accumulate the full response
                    while ((line = await reader.ReadLineAsync()) != null)
                    {
                        fullResponse.Append(line);
                    }

                    // Append the AI response to chat history after all chunks are received
                    string fullResponseString = fullResponse.ToString().Trim();
                    AppendToChatHistory("AI: " + fullResponseString);
                    Debug.WriteLine($"AI Response: {fullResponseString}");
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
                imageContent.Headers.ContentType = new MediaTypeHeaderValue("image/jpeg"); // Adjust for the file type
                content.Add(imageContent, "file", Path.GetFileName(imagePath));

                // Add the prompt and model as form data
                content.Add(new StringContent("Image analysis prompt"), "prompt"); // Example prompt
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
