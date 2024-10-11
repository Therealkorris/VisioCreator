using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using OllamaSharp;  // OllamaSharp namespace

namespace VisioPlugin
{
    public class AIChatPane : UserControl
    {
        private TextBox chatInput;
        private Button sendButton;
        private RichTextBox chatHistory;
        private Visio.Application visioApplication;
        private OllamaApiClient ollamaClient;  // OllamaSharp API client
        private string selectedModel = "llama3.2:latest";  // Default model for the API

        public AIChatPane(Visio.Application visioApp, OllamaApiClient ollamaClient)
        {
            visioApplication = visioApp;
            this.ollamaClient = ollamaClient;  // Inject Ollama API client
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            chatHistory = new RichTextBox
            {
                Dock = DockStyle.Top,
                ReadOnly = true,
                Height = 500
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
        }

        private async void SendButton_Click(object sender, EventArgs e)
        {
            string userMessage = chatInput.Text.Trim();
            if (!string.IsNullOrEmpty(userMessage))
            {
                AppendMessage("User", userMessage);
                chatInput.Clear();

                // Fetch AI response from Ollama API
                string aiResponse = await GetAIResponse(userMessage);
                AppendMessage("AI", aiResponse);
            }
        }

        // Method to get AI response from Ollama API
        private async Task<string> GetAIResponse(string userMessage)
        {
            string aiResponse = "";
            try
            {
                // Use OllamaSharp to send the user message and stream the AI's response
                await foreach (var stream in ollamaClient.Generate(userMessage))
                {
                    aiResponse += stream.Response;
                }
            }
            catch (Exception ex)
            {
                aiResponse = $"Error getting AI response: {ex.Message}";
            }

            return aiResponse;
        }

        private void AppendMessage(string sender, string message)
        {
            chatHistory.AppendText($"{sender}: {message}\n\n");
            chatHistory.ScrollToCaret();
        }
    }
}
