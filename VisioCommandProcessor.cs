using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioPlugin
{
    public class VisioCommandProcessor
    {
        private readonly Visio.Application visioApplication;
        private readonly LibraryManager libraryManager;
        private static readonly HttpClient httpClient = new HttpClient() { Timeout = TimeSpan.FromMinutes(30) };

        public VisioCommandProcessor(Visio.Application visioApp, LibraryManager libraryManager)
        {
            visioApplication = visioApp ?? throw new ArgumentNullException(nameof(visioApp));
            this.libraryManager = libraryManager ?? throw new ArgumentNullException(nameof(libraryManager));
        }

        public void ProcessCommand(string jsonCommand)
        {
            try
            {
                Debug.WriteLine($"[ProcessCommand] Received command: {jsonCommand}");
                JObject commandObject = JsonConvert.DeserializeObject<JObject>(jsonCommand);

                string commandType = commandObject["command"]?.ToString();

                if (string.IsNullOrEmpty(commandType))
                {
                    Debug.WriteLine($"[ProcessCommand] [Error] Unknown or missing command type.");
                    return;
                }

                if (commandType.Equals("CreateShapes", StringComparison.OrdinalIgnoreCase))
                {
                    commandType = "CreateShape";
                }

                // Handle different command types
                if (commandType == "CreateShape")
                {
                    if (commandObject["parameters"]?["shapes"] is JArray shapesArray)
                    {
                        foreach (JObject shapeObject in shapesArray)
                        {
                            ExecuteCreateShapeCommand(shapeObject);
                        }
                    }
                    else if (commandObject["parameters"] is JObject shapeParameters)
                    {
                        ExecuteCreateShapeCommand(shapeParameters);
                    }
                }
                else if (commandType == "ConnectShapes")
                {
                    ExecuteConnectShapesCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "AddTextToShape")
                {
                    ExecuteAddTextToShapeCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "SetShapeStyle")
                {
                    ExecuteSetShapeStyleCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "GroupShapes")
                {
                    ExecuteGroupShapesCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "UngroupShapes")
                {
                    ExecuteUngroupShapesCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "AlignShapes")
                {
                    ExecuteAlignShapesCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "DistributeShapes")
                {
                    ExecuteDistributeShapesCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "GetShapeProperties")
                {
                    ExecuteGetShapePropertiesCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "GetPageSize")
                {
                    ExecuteGetPageSizeCommand(commandObject["parameters"] as JObject);
                }
                else if (commandType == "CreateTextBox")
                {
                    ExecuteCreateTextBoxCommand(commandObject["parameters"] as JObject);
                }
                else
                {
                    Debug.WriteLine($"[ProcessCommand] [Error] Unsupported command type: {commandType}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessCommand] [Error] Failed to process command: {ex.Message}");
            }
        }

        private void ExecuteCreateShapeCommand(JObject shapeParameters)
        {
            if (shapeParameters == null)
            {
                Debug.WriteLine("[ExecuteCreateShapeCommand] [Error] Shape parameters are missing.");
                return;
            }

            // Get the active page from the Visio application
            var activePage = visioApplication.ActivePage;
            if (activePage == null)
            {
                Debug.WriteLine("[ExecuteCreateShapeCommand] [Error] No active page found.");
                return;
            }

            // Get the page dimensions
            double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
            double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

            // Check if the 'shapes' array exists
            if (shapeParameters["shapes"] is JArray shapesArray)
            {
                foreach (JObject shapeObject in shapesArray)
                {
                    CreateSingleShape(shapeObject, activePage, pageWidth, pageHeight);
                }
            }
            // Handle single shape creation
            else
            {
                CreateSingleShape(shapeParameters, activePage, pageWidth, pageHeight);
            }
        }

        private void CreateSingleShape(JObject shapeObject, Visio.Page activePage, double pageWidth, double pageHeight)
        {
            string shapeType = shapeObject["type"]?.ToString() ?? shapeObject["shapeType"]?.ToString();
            if (string.IsNullOrEmpty(shapeType))
            {
                Debug.WriteLine("[CreateSingleShape] [Error] shapeType is missing or empty.");
                return;
            }

            JObject positionObject = shapeObject["position"] as JObject;
            double xPercent = positionObject?["x"]?.Value<double>() ?? 0;
            double yPercent = positionObject?["y"]?.Value<double>() ?? 0;

            JObject sizeObject = shapeObject["size"] as JObject;
            double widthPercent = sizeObject?["width"]?.Value<double>() ?? 10;
            double heightPercent = sizeObject?["height"]?.Value<double>() ?? 10;

            // Ensure percentages are within bounds
            xPercent = Math.Max(0, Math.Min(100, xPercent));
            yPercent = Math.Max(0, Math.Min(100, yPercent));
            widthPercent = Math.Max(0, Math.Min(100, widthPercent));
            heightPercent = Math.Max(0, Math.Min(100, heightPercent));

            // Adjust position to keep shape within canvas
            double shapeWidth = (widthPercent / 100.0) * pageWidth;
            double shapeHeight = (heightPercent / 100.0) * pageHeight;
            double x = (xPercent / 100.0) * pageWidth - shapeWidth / 2;
            double y = (yPercent / 100.0) * pageHeight - shapeHeight / 2;

            x = Math.Max(0, Math.Min(pageWidth - shapeWidth, x));
            y = Math.Max(0, Math.Min(pageHeight - shapeHeight, y));

            // Convert adjusted position back to percentage
            double adjustedXPercent = (x / pageWidth) * 100;
            double adjustedYPercent = (y / pageHeight) * 100;

            string color = shapeObject["color"]?.ToString();

            // Create the shape
            var shape = libraryManager.AddShapeToDocument(libraryManager.GetCategories().FirstOrDefault(), shapeType, adjustedXPercent, adjustedYPercent, widthPercent, heightPercent);

            if (shape != null)
            {
                Debug.WriteLine($"[CreateSingleShape] Created shape of type {shapeType} with ID: {shape.ID16}");
                if (!string.IsNullOrEmpty(color))
                {
                    libraryManager.SetShapeColor(shape, color);
                }
            }
            else
            {
                Debug.WriteLine("[CreateSingleShape] Failed to create shape.");
            }
        }

        private void ExecuteConnectShapesCommand(JObject parameters)
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
        }

        private void ExecuteAddTextToShapeCommand(JObject parameters)
        {
            string shapeName = parameters?["shapeName"]?.ToString();
            string text = parameters?["text"]?.ToString();

            if (string.IsNullOrEmpty(shapeName) || string.IsNullOrEmpty(text))
            {
                Debug.WriteLine("[ExecuteAddTextToShapeCommand] [Error] shapeName or text is missing.");
                return;
            }

            libraryManager.AddTextToShape(shapeName, text);
        }

        private void ExecuteSetShapeStyleCommand(JObject parameters)
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
        }

        private void ExecuteGroupShapesCommand(JObject parameters)
        {
            var shapeNames = parameters?["shapeNames"]?.ToObject<string[]>();

            if (shapeNames == null || shapeNames.Length == 0)
            {
                Debug.WriteLine("[ExecuteGroupShapesCommand] [Error] shapeNames is missing or empty.");
                return;
            }

            libraryManager.GroupShapes(shapeNames);
        }

        private void ExecuteUngroupShapesCommand(JObject parameters)
        {
            string shapeName = parameters?["shapeName"]?.ToString();

            if (string.IsNullOrEmpty(shapeName))
            {
                Debug.WriteLine("[ExecuteUngroupShapesCommand] [Error] shapeName is missing.");
                return;
            }

            libraryManager.UngroupShapes(shapeName);
        }

        private void ExecuteAlignShapesCommand(JObject parameters)
        {
            var shapeNames = parameters?["shapeNames"]?.ToObject<string[]>();
            string alignmentType = parameters?["alignmentType"]?.ToString();

            if (shapeNames == null || shapeNames.Length == 0 || string.IsNullOrEmpty(alignmentType))
            {
                Debug.WriteLine("[ExecuteAlignShapesCommand] [Error] shapeNames or alignmentType is missing.");
                return;
            }

            libraryManager.AlignShapes(shapeNames, alignmentType);
        }

        private void ExecuteDistributeShapesCommand(JObject parameters)
        {
            var shapeNames = parameters?["shapeNames"]?.ToObject<string[]>();
            string distributionType = parameters?["distributionType"]?.ToString();

            if (shapeNames == null || shapeNames.Length == 0 || string.IsNullOrEmpty(distributionType))
            {
                Debug.WriteLine("[ExecuteDistributeShapesCommand] [Error] shapeNames or distributionType is missing.");
                return;
            }

            libraryManager.DistributeShapes(shapeNames, distributionType);
        }

        private void ExecuteGetShapePropertiesCommand(JObject parameters)
        {
            string shapeName = parameters?["shapeName"]?.ToString();
            if (string.IsNullOrEmpty(shapeName))
            {
                Debug.WriteLine("[ExecuteGetShapePropertiesCommand] [Error] shapeName is missing.");
                return;
            }

            string propertiesJson = libraryManager.GetShapeProperties(shapeName);
            try
            {
                var content = new StringContent(propertiesJson, Encoding.UTF8, "application/json");
                var response = httpClient.PostAsync("http://localhost:5678/chat-agent", content).Result;
                response.EnsureSuccessStatusCode();
                Debug.WriteLine($"[ExecuteGetShapePropertiesCommand] Sent properties for shape '{shapeName}' to n8n.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExecuteGetShapePropertiesCommand] [Error] Failed to send properties to n8n: {ex.Message}");
            }
        }

        private void ExecuteGetPageSizeCommand(JObject parameters)
        {
            string pageSizeJson = libraryManager.GetPageSize();
            try
            {
                var content = new StringContent(pageSizeJson, Encoding.UTF8, "application/json");
                var response = httpClient.PostAsync("http://localhost:5680/chat-agent", content).Result;
                response.EnsureSuccessStatusCode();
                Debug.WriteLine($"[ExecuteGetPageSizeCommand] Sent page size to n8n.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExecuteGetPageSizeCommand] [Error] Failed to send page size to n8n: {ex.Message}");
            }
        }

        private void ExecuteCreateTextBoxCommand(JObject parameters)
        {
            string content = parameters["content"]?.ToString();
            JObject position = parameters["position"] as JObject;
            double xPercent = position?["x"]?.Value<double>() ?? 0;
            double yPercent = position?["y"]?.Value<double>() ?? 0;
            double fontSize = parameters["fontSize"]?.Value<double>() ?? 12;
            string color = parameters["color"]?.ToString() ?? "black";

            if (string.IsNullOrEmpty(content) || position == null)
            {
                Debug.WriteLine("[ExecuteCreateTextBoxCommand] [Error] Missing content or position.");
                return;
            }

            var activePage = visioApplication.ActivePage;
            if (activePage == null)
            {
                Debug.WriteLine("[ExecuteCreateTextBoxCommand] [Error] No active page found.");
                return;
            }

            try
            {
                double smallWidth = 0.01;
                double smallHeight = 0.01;

                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                double visioX = (xPercent / 100.0) * pageWidth;
                double visioY = ((100 - yPercent) / 100.0) * pageHeight;

                var textShape = activePage.DrawRectangle(visioX - smallWidth / 2, visioY - smallHeight / 2, visioX + smallWidth / 2, visioY + smallHeight / 2);
                textShape.Text = content;
                textShape.CellsU["Char.Size"].FormulaU = fontSize.ToString();
                textShape.CellsU["Char.Color"].FormulaU = $"RGB({ConvertColorToRGB(color)})";

                Debug.WriteLine($"[ExecuteCreateTextBoxCommand] Added text box: '{content}' at ({visioX}, {visioY}).");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExecuteCreateTextBoxCommand] [Error] Failed to create text box: {ex.Message}");
            }
        }

        private int ConvertColorToRGB(string colorName)
        {
            return colorName.ToLower() switch
            {
                "black" => 0,
                "red" => 255,
                "green" => 65280,
                "blue" => 16711680,
                _ => 0 // Default to black
            };
        }
    }
}