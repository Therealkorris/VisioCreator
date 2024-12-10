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

                // Handle empty or invalid commands
                if (string.IsNullOrEmpty(commandType))
                {
                    Debug.WriteLine($"[ProcessCommand] [Error] Unknown or missing command type.");
                    return;
                }

                // Map command variations to the correct command type
                if (commandType.Equals("CreateShapes", StringComparison.OrdinalIgnoreCase))
                {
                    commandType = "CreateShape"; // Correct the command type
                }

                // Handle different command types
                if (commandType == "CreateShape")
                {
                    // Check for the shapes array (multiple shapes)
                    if (commandObject["parameters"]?["shapes"] is JArray shapesArray)
                    {
                        foreach (JObject shapeObject in shapesArray)
                        {
                            await ExecuteCreateShapeCommand(shapeObject);
                        }
                    }
                    // Handle case where parameters are directly in 'parameters' object (single shape)
                    else if (commandObject["parameters"] is JObject shapeParameters)
                    {
                        await ExecuteCreateShapeCommand(shapeParameters);
                    }
                    else
                    {
                        Debug.WriteLine("[ProcessCommand] [Error] 'parameters' is missing or has an invalid format.");
                        return;
                    }
                }
                // Add other command types here (e.g., ConnectShapes, AddTextToShape, etc.)
                else if (commandType == "ConnectShapes")
                {
                    await ExecuteConnectShapesCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "AddTextToShape")
                {
                    await ExecuteAddTextToShapeCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "SetShapeStyle")
                {
                    await ExecuteSetShapeStyleCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "GroupShapes")
                {
                    await ExecuteGroupShapesCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "UngroupShapes")
                {
                    await ExecuteUngroupShapesCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "AlignShapes")
                {
                    await ExecuteAlignShapesCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "DistributeShapes")
                {
                    await ExecuteDistributeShapesCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "GetShapeProperties")
                {
                    await ExecuteGetShapePropertiesCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "GetPageSize")
                {
                    await ExecuteGetPageSizeCommand(commandObject["parameters"] as JObject);
                }
                else
                {
                    Debug.WriteLine($"[ProcessCommand] [Error] Unsupported command type: {commandType}");
                }
            }
            catch (JsonReaderException jEx)
            {
                Debug.WriteLine($"[ProcessCommand] [Error] Invalid JSON format: {jEx.Message}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessCommand] [Error] Failed to process command: {ex.Message}");
            }
        }

        private async Task ExecuteCreateShapeCommand(JObject shapeParameters)
        {
            if (shapeParameters == null)
            {
                Debug.WriteLine("[ExecuteCreateShapeCommand] [Error] Shape parameters are missing.");
                return;
            }

            // Check if the 'shapes' array exists
            if (shapeParameters["shapes"] is JArray shapesArray)
            {
                foreach (JObject shapeObject in shapesArray)
                {
                    string shapeType = shapeObject["type"]?.ToString();
                    if (string.IsNullOrEmpty(shapeType))
                    {
                        Debug.WriteLine("[ExecuteCreateShapeCommand] [Error] shapeType is missing or empty in one of the shape objects.");
                        continue; // Skip this shape object and move to the next one
                    }

                    JObject positionObject = shapeObject["position"] as JObject;
                    double x = positionObject?["x"]?.Value<double>() ?? 0;
                    double y = positionObject?["y"]?.Value<double>() ?? 0;

                    JObject sizeObject = shapeObject["size"] as JObject;
                    double width = sizeObject?["width"]?.Value<double>() ?? 10;
                    double height = sizeObject?["height"]?.Value<double>() ?? 10;

                    string color = shapeObject["color"]?.ToString();

                    libraryManager.AddShapeToDocument(libraryManager.GetCategories().FirstOrDefault(), shapeType, x, y, width, height);

                    if (!string.IsNullOrEmpty(color))
                    {
                        var shapeName = GetLastAddedShapeName();
                        if (!string.IsNullOrEmpty(shapeName))
                        {
                            var shape = visioApplication.ActivePage.Shapes.ItemU[shapeName];
                            libraryManager.SetShapeColor(shape, color);
                        }
                    }
                }
            }
            // Handle the case where 'shapes' array is missing but other parameters are present (single shape creation)
            else
            {
                string shapeType = shapeParameters["shapeType"]?.ToString();
                if (string.IsNullOrEmpty(shapeType))
                {
                    Debug.WriteLine("[ExecuteCreateShapeCommand] [Error] shapeType is missing or empty.");
                    return;
                }

                JObject positionObject = shapeParameters["position"] as JObject;
                double x = positionObject?["x"]?.Value<double>() ?? 0;
                double y = positionObject?["y"]?.Value<double>() ?? 0;

                JObject sizeObject = shapeParameters["size"] as JObject;
                double width = sizeObject?["width"]?.Value<double>() ?? 10;
                double height = sizeObject?["height"]?.Value<double>() ?? 10;

                string color = shapeParameters["color"]?.ToString();

                libraryManager.AddShapeToDocument(libraryManager.GetCategories().FirstOrDefault(), shapeType, x, y, width, height);

                if (!string.IsNullOrEmpty(color))
                {
                    var shapeName = GetLastAddedShapeName();
                    if (!string.IsNullOrEmpty(shapeName))
                    {
                        var shape = visioApplication.ActivePage.Shapes.ItemU[shapeName];
                        libraryManager.SetShapeColor(shape, color);
                    }
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
                    var response = await client.PostAsync("http://localhost:5678/chat-agent", content); // Replace with your n8n webhook URL
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