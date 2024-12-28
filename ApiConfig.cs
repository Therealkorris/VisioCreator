using System;
using System.Diagnostics;
using System.Linq;

namespace VisioPlugin
{
    public static class ApiConfig
    {
        private static string _baseUrl = "http://localhost:5678/webhook";
        public static string BaseUrl => _baseUrl;
        public static bool IsTestEnvironment { get; private set; }
        public static int N8nPort => 5678;
        public static int VisioPort => 5680;

        public static void UpdateFromApiEndpoint(string endpoint)
        {
            if (string.IsNullOrEmpty(endpoint))
                throw new ArgumentNullException(nameof(endpoint));

            if (!Uri.TryCreate(endpoint, UriKind.Absolute, out Uri uri))
                throw new ArgumentException("Invalid URL format", nameof(endpoint));

            _baseUrl = endpoint;
            IsTestEnvironment = endpoint.Contains("webhook-test");
            Debug.WriteLine($"[ApiConfig] Updated base URL to: {_baseUrl}");
            Debug.WriteLine($"[ApiConfig] Test environment: {IsTestEnvironment}");
        }

        public static string GetWebhookUrl(string endpoint = "")
        {
            // Always use N8nPort (5678) for outgoing requests to the backend
            var baseUri = new Uri(_baseUrl);
            var path = baseUri.AbsolutePath.TrimEnd('/');
            var url = $"http://localhost:{N8nPort}{path}/{endpoint}".TrimEnd('/');
            Debug.WriteLine($"[ApiConfig] Generated webhook URL: {url}");
            return url;
        }

        public static string GetVisioWebhookUrl(string endpoint = "")
        {
            // Use direct paths for all Visio endpoints
            var url = $"http://localhost:{VisioPort}/{endpoint}/".TrimEnd('/');
            Debug.WriteLine($"[ApiConfig] Generated Visio URL: {url}");
            return url;
        }
    }
} 