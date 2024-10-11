using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using OllamaSharp;  // OllamaSharp namespace
using System.Diagnostics;

namespace VisioPlugin
{
    public class AIChatPane : Form  // Change from UserControl to Form
    {
        private TextBox chatInput;
        private Button sendButton;
        private RichTextBox chatHistory;
        private Visio.Application visioApplication;
        private OllamaApiClient ollamaClient;  // OllamaSharp API client

        public AIChatPane(Visio.Application visioApp, OllamaApiClient ollamaClient)
        {
            visioApplication = visioApp;
            this.ollamaClient = ollamaClient;  // Inject Ollama API client
            InitializeComponent();
            
            // Set form properties
            this.Text = "AI Chat";
            this.Size = new System.Drawing.Size(400, 600);

            Debug.WriteLine("AIChatPane initialized");
        }

        private void InitializeComponent()
        {
            try
            {
                chatHistory = new RichTextBox
                {
                    Dock = DockStyle.Fill,
                    ReadOnly = true
                };

                chatInput = new TextBox
                {
                    Dock = DockStyle.Bottom,
                    Height = 50
                };

                sendButton = new Button
                {
                    Text = "Send",
                    Dock = DockStyle.Bottom,
                    Height = 30
                };
                sendButton.Click += SendButton_Click;

                Controls.Add(chatHistory);
                Controls.Add(chatInput);
                Controls.Add(sendButton);

                // Set the form's layout
                this.Layout += (sender, e) => PerformLayout();

                Debug.WriteLine("AIChatPane components initialized");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in InitializeComponent: {ex.Message}");
            }
        }

        private async void SendButton_Click(object sender, EventArgs e)
        {
            string userMessage = chatInput.Text.Trim();
            if (!string.IsNullOrEmpty(userMessage))
            {
                AppendMessage("User", userMessage);
                chatInput.Clear();

                sendButton.Enabled = false;
                string aiResponse = await GetAIResponse(userMessage);
                AppendMessage("AI", aiResponse);
                sendButton.Enabled = true;
            }
        }

        // Method to get AI response from Ollama API
        private async Task<string> GetAIResponse(string userMessage)
        {
            string aiResponse = "";
            try
            {
                Debug.WriteLine($"Sending message to AI: {userMessage}");
                await foreach (var stream in ollamaClient.Generate(userMessage))
                {
                    aiResponse += stream.Response;
                    Debug.WriteLine($"Received partial response: {stream.Response}");
                }
                Debug.WriteLine($"Full AI response: {aiResponse}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting AI response: {ex.Message}");
                aiResponse = $"Error getting AI response: {ex.Message}";
            }

            return aiResponse;
        }

        private void AppendMessage(string sender, string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => AppendMessage(sender, message)));
                return;
            }

            chatHistory.AppendText($"{sender}: {message}\n\n");
            chatHistory.ScrollToCaret();
        }
    }
}
