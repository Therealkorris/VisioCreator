using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioPlugin
{
    public class VisioCommandProcessor
    {
        private readonly Visio.Application visioApplication;
        private readonly LibraryManager libraryManager;

        public VisioCommandProcessor(Visio.Application visioApp, LibraryManager libraryManager)
        {
            visioApplication = visioApp ?? throw new ArgumentNullException(nameof(visioApp));
            this.libraryManager = libraryManager ?? throw new ArgumentNullException(nameof(libraryManager));
        }

        public async Task ProcessCommand(string jsonCommand)
        {
            try
            {
                Debug.WriteLine($"[ProcessCommand] Received command: {jsonCommand}");
                JObject commandObject = JsonConvert.DeserializeObject<JObject>(jsonCommand);

                string commandType = commandObject["command"]?.ToString();
                JObject parameters = commandObject["parameters"] as JObject;

                if (string.IsNullOrEmpty(commandType))
                {
                    Debug.WriteLine("[ProcessCommand] [Error] Command type is missing or empty.");
                    return;
                }

                switch (commandType)
                {
                    case "CreateShape":
                        await ExecuteCreateShapeCommand(parameters);
                        break;
                    case "ConnectShapes":
                        await ExecuteConnectShapesCommand(parameters);
                        break;
                    case "AddTextToShape":
                        await ExecuteAddTextToShapeCommand(parameters);
                        break;
                    case "SetShapeStyle":
                        await ExecuteSetShapeStyleCommand(parameters);
                        break;
                    case "GroupShapes":
                        await ExecuteGroupShapesCommand(parameters);
                        break;
                    case "UngroupShapes":
                        await ExecuteUngroupShapesCommand(parameters);
                        break;
                    case "AlignShapes":
                        await ExecuteAlignShapesCommand(parameters);
                        break;
                    case "DistributeShapes":
                        await ExecuteDistributeShapesCommand(parameters);
                        break;
                    case "GetShapeProperties":
                        await ExecuteGetShapePropertiesCommand(parameters);
                        break;
                    case "GetPageSize":
                        await ExecuteGetPageSizeCommand(parameters);
                        break;
                    default:
                        Debug.WriteLine($"[ProcessCommand] [Error] Unknown command type: {commandType}");
                        break;
                }
            }
            catch (JsonReaderException jEx)
            {
                Debug.WriteLine($"[ProcessCommand] [Error] Invalid JSON format: {jEx.Message}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessCommand] [Error] Failed to process command: {ex.Message}");
                // Consider sending detailed error information back to the AI for debugging.
            }
        }

        private async Task ExecuteCreateShapeCommand(JObject parameters)
        {
            if (parameters == null)
            {
                Debug.WriteLine("[ExecuteCreateShapeCommand] [Error] Parameters are missing.");
                return;
            }

            // Extract parameters with more robust error handling
            string shapeType = parameters["shapeType"]?.ToString();
            if (string.IsNullOrEmpty(shapeType))
            {
                Debug.WriteLine("[ExecuteCreateShapeCommand] [Error] shapeType is missing or empty.");
                return; // Consider sending an error message back to the AI
            }

            JObject positionObject = parameters["position"] as JObject;
            double x = positionObject?["x"]?.Value<double>() ?? 0;
            double y = positionObject?["y"]?.Value<double>() ?? 0;

            JObject sizeObject = parameters["size"] as JObject;
            double width = sizeObject?["width"]?.Value<double>() ?? 10; // Default width
            double height = sizeObject?["height"]?.Value<double>() ?? 10; // Default height

            string color = parameters["color"]?.ToString();

            // Use the parameters to add the shape
            libraryManager.AddShapeToDocument(libraryManager.GetCategories().FirstOrDefault(), shapeType, x, y, width, height);

            // Optionally, set the shape color if provided
            if (!string.IsNullOrEmpty(color))
            {
                var shapeName = GetLastAddedShapeName();
                if (!string.IsNullOrEmpty(shapeName))
                {
                    var shape = visioApplication.ActivePage.Shapes.ItemU[shapeName];
                    libraryManager.SetShapeColor(shape, color);
                }
            }

            await Task.CompletedTask;
        }

        private string GetLastAddedShapeName()
        {
            try
            {
                var activePage = visioApplication.ActivePage;
                if (activePage != null && activePage.Shapes.Count > 0)
                {
                    // Assuming the last added shape is at the end of the Shapes collection
                    return activePage.Shapes[activePage.Shapes.Count].Name;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[GetLastAddedShapeName] Error: {ex.Message}");
            }

            return null;
        }

        private async Task ExecuteConnectShapesCommand(JObject parameters)
        {
            string shape1Name = parameters?["shape1Name"]?.ToString();
            string shape2Name = parameters?["shape2Name"]?.ToString();
            string connectorType = parameters?["connectorType"]?.ToString();

            if (string.IsNullOrEmpty(shape1Name) || string.IsNullOrEmpty(shape2Name))
            {
                Debug.WriteLine("[ExecuteConnectShapesCommand] [Error] shape1Name or shape2Name is missing.");
                return;
            }

            libraryManager.ConnectShapes(shape1Name, shape2Name, connectorType);
            await Task.CompletedTask;
        }

        private async Task ExecuteAddTextToShapeCommand(JObject parameters)
        {
            string shapeName = parameters?["shapeName"]?.ToString();
            string text = parameters?["text"]?.ToString();

            if (string.IsNullOrEmpty(shapeName) || string.IsNullOrEmpty(text))
            {
                Debug.WriteLine("[ExecuteAddTextToShapeCommand] [Error] shapeName or text is missing.");
                return;
            }

            libraryManager.AddTextToShape(shapeName, text);
            await Task.CompletedTask;
        }

        private async Task ExecuteSetShapeStyleCommand(JObject parameters)
        {
            string shapeName = parameters?["shapeName"]?.ToString();
            string lineStyle = parameters?["lineStyle"]?.ToString();
            string fillPattern = parameters?["fillPattern"]?.ToString();

            if (string.IsNullOrEmpty(shapeName))
            {
                Debug.WriteLine("[ExecuteSetShapeStyleCommand] [Error] shapeName is missing.");
                return;
            }

            libraryManager.SetShapeStyle(shapeName, lineStyle, fillPattern);
            await Task.CompletedTask;
        }

        private async Task ExecuteGroupShapesCommand(JObject parameters)
        {
            var shapeNames = parameters?["shapeNames"]?.ToObject<string[]>();

            if (shapeNames == null || shapeNames.Length == 0)
            {
                Debug.WriteLine("[ExecuteGroupShapesCommand] [Error] shapeNames is missing or empty.");
                return;
            }

            libraryManager.GroupShapes(shapeNames);
            await Task.CompletedTask;
        }

        private async Task ExecuteUngroupShapesCommand(JObject parameters)
        {
            string shapeName = parameters?["shapeName"]?.ToString();

            if (string.IsNullOrEmpty(shapeName))
            {
                Debug.WriteLine("[ExecuteUngroupShapesCommand] [Error] shapeName is missing.");
                return;
            }

            libraryManager.UngroupShapes(shapeName);
            await Task.CompletedTask;
        }

        private async Task ExecuteAlignShapesCommand(JObject parameters)
        {
            var shapeNames = parameters?["shapeNames"]?.ToObject<string[]>();
            string alignmentType = parameters?["alignmentType"]?.ToString();

            if (shapeNames == null || shapeNames.Length == 0 || string.IsNullOrEmpty(alignmentType))
            {
                Debug.WriteLine("[ExecuteAlignShapesCommand] [Error] shapeNames or alignmentType is missing.");
                return;
            }

            libraryManager.AlignShapes(shapeNames, alignmentType);
            await Task.CompletedTask;
        }

        private async Task ExecuteDistributeShapesCommand(JObject parameters)
        {
            var shapeNames = parameters?["shapeNames"]?.ToObject<string[]>();
            string distributionType = parameters?["distributionType"]?.ToString();

            if (shapeNames == null || shapeNames.Length == 0 || string.IsNullOrEmpty(distributionType))
            {
                Debug.WriteLine("[ExecuteDistributeShapesCommand] [Error] shapeNames or distributionType is missing.");
                return;
            }

            libraryManager.DistributeShapes(shapeNames, distributionType);
            await Task.CompletedTask;
        }

        private async Task ExecuteGetShapePropertiesCommand(JObject parameters)
        {
            string shapeName = parameters?["shapeName"]?.ToString();
            if (string.IsNullOrEmpty(shapeName))
            {
                Debug.WriteLine("[ExecuteGetShapePropertiesCommand] [Error] shapeName is missing.");
                return;
            }

            // Get the shape properties
            string propertiesJson = libraryManager.GetShapeProperties(shapeName);

            // Send the properties back to the AI (via n8n)
            // You'll need to set up an HTTP client to send the data back to your n8n webhook
            using (var client = new HttpClient())
            {
                try
                {
                    var content = new StringContent(propertiesJson, Encoding.UTF8, "application/json");
                    var response = await client.PostAsync("http://localhost:5680/chat-agent", content); // Replace with your n8n webhook URL
                    response.EnsureSuccessStatusCode();
                    Debug.WriteLine($"[ExecuteGetShapePropertiesCommand] Sent properties for shape '{shapeName}' to n8n.");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ExecuteGetShapePropertiesCommand] [Error] Failed to send properties to n8n: {ex.Message}");
                }
            }
        }

        private async Task ExecuteGetPageSizeCommand(JObject parameters)
        {
            // Get the page size
            string pageSizeJson = libraryManager.GetPageSize();

            // Send the page size back to the AI (via n8n)
            using (var client = new HttpClient())
            {
                try
                {
                    var content = new StringContent(pageSizeJson, Encoding.UTF8, "application/json");
                    var response = await client.PostAsync("http://localhost:5680/chat-agent", content); // Replace with your n8n webhook URL
                    response.EnsureSuccessStatusCode();
                    Debug.WriteLine($"[ExecuteGetPageSizeCommand] Sent page size to n8n.");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ExecuteGetPageSizeCommand] [Error] Failed to send page size to n8n: {ex.Message}");
                }
            }
        }
    }
}