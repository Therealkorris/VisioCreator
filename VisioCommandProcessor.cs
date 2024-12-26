using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
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
        private static readonly HttpClient httpClient = new HttpClient() { Timeout = TimeSpan.FromMinutes(30) };
        private List<Visio.Shape> lastAffectedShapes = new List<Visio.Shape>();

        public VisioCommandProcessor(Visio.Application visioApp, LibraryManager libraryManager)
        {
            visioApplication = visioApp ?? throw new ArgumentNullException(nameof(visioApp));
            this.libraryManager = libraryManager ?? throw new ArgumentNullException(nameof(libraryManager));
        }

        public IEnumerable<Visio.Shape> GetLastAffectedShapes()
        {
            Debug.WriteLine($"[GetLastAffectedShapes] Returning {lastAffectedShapes.Count} shapes:");
            foreach (var shape in lastAffectedShapes)
            {
                Debug.WriteLine($"[GetLastAffectedShapes] Shape ID: {shape.ID16}");
            }
            return new List<Visio.Shape>(lastAffectedShapes);
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

                // Clear the affected shapes list before processing new command
                lastAffectedShapes = new List<Visio.Shape>();
                Debug.WriteLine("[ProcessCommand] Cleared affected shapes list.");

                // Handle different command types
                if (commandType == "CreateShape")
                {
                    // Check for the shapes array (multiple shapes)
                    if (commandObject["parameters"]?["shapes"] is JArray shapesArray)
                    {
                        foreach (JObject shapeObject in shapesArray)
                        {
                            var shape = ExecuteCreateShapeCommand(shapeObject);
                            if (shape != null && !lastAffectedShapes.Any(s => s.ID16 == shape.ID16))
                            {
                                lastAffectedShapes.Add(shape);
                                Debug.WriteLine($"[ProcessCommand] Added shape to affected shapes. ID: {shape.ID16}");
                            }
                        }
                    }
                    // Handle case where parameters are directly in 'parameters' object (single shape)
                    else if (commandObject["parameters"] is JObject shapeParameters)
                    {
                        var shape = ExecuteCreateShapeCommand(shapeParameters);
                        if (shape != null && !lastAffectedShapes.Any(s => s.ID16 == shape.ID16))
                        {
                            lastAffectedShapes.Add(shape);
                            Debug.WriteLine($"[ProcessCommand] Added single shape to affected shapes. ID: {shape.ID16}");
                        }
                    }
                    else
                    {
                        Debug.WriteLine("[ProcessCommand] [Error] 'parameters' is missing or has an invalid format.");
                        return;
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

                // Debug output for affected shapes
                Debug.WriteLine($"[ProcessCommand] Final number of affected shapes: {lastAffectedShapes.Count}");
                foreach (var shape in lastAffectedShapes.Distinct())
                {
                    Debug.WriteLine($"[ProcessCommand] Final affected shape ID: {shape.ID16}");
                }
            }
            catch (JsonReaderException jEx)
            {
                Debug.WriteLine($"[ProcessCommand] [Error] Invalid JSON format: {jEx.Message}");
                lastAffectedShapes.Clear();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessCommand] [Error] Failed to process command: {ex.Message}");
                lastAffectedShapes.Clear();
            }
        }

        private Visio.Shape ExecuteCreateShapeCommand(JObject shapeParameters)
        {
            if (shapeParameters == null)
            {
                Debug.WriteLine("[ExecuteCreateShapeCommand] [Error] Shape parameters are missing.");
                return null;
            }

            // Get the active page from the Visio application
            var activePage = visioApplication.ActivePage;
            if (activePage == null)
            {
                Debug.WriteLine("[ExecuteCreateShapeCommand] [Error] No active page found.");
                return null;
            }

            // Get the page dimensions
            double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
            double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

            // Check if the 'shapes' array exists
            if (shapeParameters["shapes"] is JArray shapesArray)
            {
                Visio.Shape lastShape = null;
                foreach (JObject shapeObject in shapesArray)
                {
                    var shape = CreateSingleShape(shapeObject, activePage, pageWidth, pageHeight);
                    if (shape != null && !lastAffectedShapes.Any(s => s.ID16 == shape.ID16))
                    {
                        lastAffectedShapes.Add(shape);
                        lastShape = shape;
                        Debug.WriteLine($"[ExecuteCreateShapeCommand] Added shape to affected shapes. ID: {shape.ID16}");
                    }
                }
                return lastShape;
            }
            // Handle single shape creation
            else
            {
                var shape = CreateSingleShape(shapeParameters, activePage, pageWidth, pageHeight);
                if (shape != null && !lastAffectedShapes.Any(s => s.ID16 == shape.ID16))
                {
                    lastAffectedShapes.Add(shape);
                    Debug.WriteLine($"[ExecuteCreateShapeCommand] Added single shape to affected shapes. ID: {shape.ID16}");
                }
                return shape;
            }
        }

        private Visio.Shape CreateSingleShape(JObject shapeObject, Visio.Page activePage, double pageWidth, double pageHeight)
        {
            string shapeType = shapeObject["type"]?.ToString() ?? shapeObject["shapeType"]?.ToString();
            if (string.IsNullOrEmpty(shapeType))
            {
                Debug.WriteLine("[CreateSingleShape] [Error] shapeType is missing or empty.");
                return null;
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

            // Get the created shape directly from AddShapeToDocument
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

            return shape;
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

            var activePage = visioApplication.ActivePage;
            if (activePage != null)
            {
                var shape1 = activePage.Shapes.ItemU[shape1Name];
                var shape2 = activePage.Shapes.ItemU[shape2Name];
                if (shape1 != null && shape2 != null)
                {
                    var connectorShape = libraryManager.ConnectShapes(shape1Name, shape2Name, connectorType);
                    if (connectorShape != null)
                    {
                        lastAffectedShapes.Add(shape1);
                        lastAffectedShapes.Add(shape2);
                        lastAffectedShapes.Add(connectorShape);
                    }
                }
            }
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

            var shape = libraryManager.AddTextToShape(shapeName, text);
            if (shape != null)
            {
                lastAffectedShapes.Add(shape);
            }
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

            var shape = libraryManager.SetShapeStyle(shapeName, lineStyle, fillPattern);
            if (shape != null)
            {
                lastAffectedShapes.Add(shape);
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
                // Create a tiny rectangle to host the text
                double smallWidth = 0.01; // Very small width
                double smallHeight = 0.01; // Very small height

                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                // Calculate coordinates in Visio units
                double visioX = (xPercent / 100.0) * pageWidth;
                double visioY = ((100 - yPercent) / 100.0) * pageHeight; // Visio Y-axis is inverted

                // Draw a small rectangle
                var textShape = activePage.DrawRectangle(visioX - smallWidth / 2, visioY - smallHeight / 2, visioX + smallWidth / 2, visioY + smallHeight / 2);

                // Add text to the rectangle
                textShape.Text = content;

                // Set font size and color
                textShape.CellsU["Char.Size"].FormulaU = fontSize.ToString();
                textShape.CellsU["Char.Color"].FormulaU = $"RGB({ConvertColorToRGB(color)})";

                // Track the affected shape
                lastAffectedShapes.Add(textShape);

                Debug.WriteLine($"[ExecuteCreateTextBoxCommand] Added text box: '{content}' at ({visioX}, {visioY}).");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExecuteCreateTextBoxCommand] [Error] Failed to create text box: {ex.Message}");
            }
        }

        // Helper method to convert color names to RGB values
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

        private void ExecuteGroupShapesCommand(JObject parameters)
        {
            var shapeNames = parameters?["shapeNames"]?.ToObject<string[]>();

            if (shapeNames == null || shapeNames.Length == 0)
            {
                Debug.WriteLine("[ExecuteGroupShapesCommand] [Error] shapeNames is missing or empty.");
                return;
            }

            var groupedShape = libraryManager.GroupShapes(shapeNames);
            if (groupedShape != null)
            {
                lastAffectedShapes.Add(groupedShape);
            }
        }

        private void ExecuteUngroupShapesCommand(JObject parameters)
        {
            string shapeName = parameters?["shapeName"]?.ToString();

            if (string.IsNullOrEmpty(shapeName))
            {
                Debug.WriteLine("[ExecuteUngroupShapesCommand] [Error] shapeName is missing.");
                return;
            }

            var ungroupedShapes = libraryManager.UngroupShapes(shapeName);
            if (ungroupedShapes.Any())
            {
                lastAffectedShapes.AddRange(ungroupedShapes);
            }
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

            var alignedShapes = libraryManager.AlignShapes(shapeNames, alignmentType);
            if (alignedShapes.Any())
            {
                lastAffectedShapes.AddRange(alignedShapes);
            }
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

            var distributedShapes = libraryManager.DistributeShapes(shapeNames, distributionType);
            if (distributedShapes.Any())
            {
                lastAffectedShapes.AddRange(distributedShapes);
            }
        }

        private void ExecuteGetShapePropertiesCommand(JObject parameters)
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
            try
            {
                var content = new StringContent(propertiesJson, Encoding.UTF8, "application/json");
                var response = httpClient.PostAsync("http://localhost:5678/chat-agent", content).Result; // Replace with your n8n webhook URL
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
            // Get the page size
            string pageSizeJson = libraryManager.GetPageSize();

            // Send the page size back to the AI (via n8n)
            try
            {
                var content = new StringContent(pageSizeJson, Encoding.UTF8, "application/json");
                var response = httpClient.PostAsync("http://localhost:5680/chat-agent", content).Result; // Replace with your n8n webhook URL
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