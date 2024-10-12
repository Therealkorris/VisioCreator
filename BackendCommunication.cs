using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Net.Http.Headers;

namespace VisioPlugin
{
    public static class BackendCommunication
    {
        private static readonly HttpClient httpClient = new HttpClient();

        public static async Task<string> SendTextMessage(string message, string model, string apiEndpoint)
        {
            var content = new MultipartFormDataContent();
            content.Add(new StringContent(message), "prompt");
            content.Add(new StringContent(model), "model");

            var response = await httpClient.PostAsync($"{apiEndpoint}/text-prompt", content);
            return await response.Content.ReadAsStringAsync();
        }

        public static async Task<string> SendImageMessage(string imagePath, string model, string apiEndpoint)
        {
            var content = new MultipartFormDataContent();
            content.Add(new StringContent(model), "model");

            var imageContent = new ByteArrayContent(File.ReadAllBytes(imagePath));
            imageContent.Headers.ContentType = new MediaTypeHeaderValue("image/jpeg");
            content.Add(imageContent, "file", Path.GetFileName(imagePath));

            var response = await httpClient.PostAsync($"{apiEndpoint}/image-prompt", content);
            return await response.Content.ReadAsStringAsync();
        }
    }
}
