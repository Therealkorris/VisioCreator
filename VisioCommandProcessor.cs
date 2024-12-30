using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
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
        private List<VisioPlugin.ShapeInfo> currentCommandShapes;

        public VisioCommandProcessor(Visio.Application visioApp, LibraryManager libraryManager)
        {
            visioApplication = visioApp ?? throw new ArgumentNullException(nameof(visioApp));
            this.libraryManager = libraryManager ?? throw new ArgumentNullException(nameof(libraryManager));
            currentCommandShapes = new List<VisioPlugin.ShapeInfo>();
        }

        public List<VisioPlugin.ShapeInfo> ProcessCommand(string jsonCommand)
        {
            Debug.WriteLine($"[ProcessCommand] Starting command: {jsonCommand}");
            currentCommandShapes = new List<ShapeInfo>();

            try
            {
                Debug.WriteLine($"[ProcessCommand] Clearing currentCommandShapes list");
                currentCommandShapes.Clear();

                // Try parsing as array first
                JToken parsedCommand = JToken.Parse(jsonCommand);
                
                if (parsedCommand is JArray commandArray)
                {
                    Debug.WriteLine($"[ProcessCommand] Processing array of commands");
                    foreach (JObject commandObject in commandArray)
                    {
                        ProcessSingleCommand(commandObject);
                    }
                }
                else if (parsedCommand is JObject singleCommand)
                {
                    Debug.WriteLine($"[ProcessCommand] Processing single command");
                    ProcessSingleCommand(singleCommand);
                }
                else
                {
                    Debug.WriteLine($"[ProcessCommand] [Error] Invalid command format. Expected object or array.");
                    return currentCommandShapes;
                }

                Debug.WriteLine($"[ProcessCommand] Command completed successfully. Total affected shapes: {currentCommandShapes.Count}");
                var result = new List<ShapeInfo>(currentCommandShapes);
                Debug.WriteLine($"[ProcessCommand] Returning {result.Count} shapes");
                foreach (var shape in result)
                {
                    Debug.WriteLine($"[ProcessCommand] Returning shape: {shape}");
                }
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessCommand] Error processing command: {ex.Message}");
                return currentCommandShapes;
            }
        }

        private void ProcessSingleCommand(JObject commandObject)
        {
            string commandType = commandObject["command"]?.ToString();

            if (string.IsNullOrEmpty(commandType))
            {
                Debug.WriteLine($"[ProcessSingleCommand] [Error] Unknown or missing command type.");
                return;
            }

            if (commandType.Equals("CreateShapes", StringComparison.OrdinalIgnoreCase))
            {
                commandType = "CreateShape";
            }

            // Handle different command types
            switch (commandType)
            {
                case "CreateShape":
                    Debug.WriteLine($"[ProcessSingleCommand] Processing CreateShape command");
                    if (commandObject["parameters"]?["shapes"] is JArray shapesArray)
                    {
                        // Reorder shapes - non-connectors first, connectors last
                        var orderedShapes = shapesArray
                            .Select(s => s as JObject)
                            .Where(s => s != null)
                            .OrderBy(s => s["shape_type"]?.ToString()?.ToLower()?.Contains("connector") ?? false)
                            .ToList();

                        Debug.WriteLine($"[ProcessSingleCommand] Processing {orderedShapes.Count} shapes in order:");
                        foreach (var shape in orderedShapes)
                        {
                            Debug.WriteLine($"  - {shape["shape_type"]?.ToString() ?? "unknown"} (ID: {shape["shape_id"]?.ToString() ?? "unknown"})");
                        }

                        // Track created shape IDs to prevent duplicates
                        var createdShapeIds = new HashSet<string>();
                        foreach (JObject shapeObject in orderedShapes)
                        {
                            string shapeId = shapeObject["shape_id"]?.ToString();
                            if (!string.IsNullOrEmpty(shapeId) && !createdShapeIds.Contains(shapeId))
                            {
                                ExecuteCreateShapeCommand(shapeObject);
                                createdShapeIds.Add(shapeId);
                            }
                            else
                            {
                                Debug.WriteLine($"[ProcessSingleCommand] Skipping duplicate shape with ID: {shapeId}");
                            }
                        }
                    }
                    else if (commandObject["parameters"] is JObject shapeParameters)
                    {
                        ExecuteCreateShapeCommand(shapeParameters);
                    }
                    break;
                case "ConnectShapes":
                    ExecuteConnectShapesCommand(commandObject["parameters"] as JObject);
                    break;
                case "AddTextToShape":
                    ExecuteAddTextToShapeCommand(commandObject["parameters"] as JObject);
                    break;
                case "SetShapeStyle":
                    ExecuteSetShapeStyleCommand(commandObject["parameters"] as JObject);
                    break;
                case "GroupShapes":
                    ExecuteGroupShapesCommand(commandObject["parameters"] as JObject);
                    break;
                case "UngroupShapes":
                    ExecuteUngroupShapesCommand(commandObject["parameters"] as JObject);
                    break;
                case "AlignShapes":
                    ExecuteAlignShapesCommand(commandObject["parameters"] as JObject);
                    break;
                case "DistributeShapes":
                    ExecuteDistributeShapesCommand(commandObject["parameters"] as JObject);
                    break;
                case "GetShapeProperties":
                    ExecuteGetShapePropertiesCommand(commandObject["parameters"] as JObject);
                    break;
                case "GetPageSize":
                    ExecuteGetPageSizeCommand(commandObject["parameters"] as JObject);
                    break;
                case "CreateTextBox":
                    ExecuteCreateTextBoxCommand(commandObject["parameters"] as JObject);
                    break;
                default:
                    Debug.WriteLine($"[ProcessSingleCommand] [Error] Unsupported command type: {commandType}");
                    break;
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

            // Process single shape creation
            CreateSingleShape(shapeParameters, activePage, pageWidth, pageHeight);
        }

        private void CreateSingleShape(JObject shapeObject, Visio.Page activePage, double pageWidth, double pageHeight)
        {
            try
            {
                // Extract all possible shape properties
                var shapeInfo = new ShapeInfo
                {
                    Category = shapeObject["category"]?.ToString(),
                    ShapeType = shapeObject["shape_type"]?.ToString() ?? shapeObject["type"]?.ToString(),
                    ShapeId = shapeObject["shape_id"]?.ToString(),
                    ShapeColor = shapeObject["shape_color"]?.ToString() ?? shapeObject["color"]?.ToString(),
                    Text = shapeObject["text"]?.ToString(),
                    PosX = shapeObject["pos_x"]?.Value<double>() ?? shapeObject["position"]?["x"]?.Value<double>() ?? 0,
                    PosY = shapeObject["pos_y"]?.Value<double>() ?? shapeObject["position"]?["y"]?.Value<double>() ?? 0,
                    Width = shapeObject["width"]?.Value<double>() ?? shapeObject["size"]?["width"]?.Value<double>() ?? 10,
                    Height = shapeObject["height"]?.Value<double>() ?? shapeObject["size"]?["height"]?.Value<double>() ?? 10,
                    Angle = shapeObject["angle"]?.Value<double>() ?? 0,
                    ZOrder = shapeObject["z_order"]?.Value<int>() ?? 0,
                    IsConnector = shapeObject["is_connector"]?.Value<bool>() ?? false,
                    ConnectorType = shapeObject["connector_type"]?.ToString(),
                    SourceShapeId = shapeObject["source_shape_id"]?.ToString(),
                    TargetShapeId = shapeObject["target_shape_id"]?.ToString(),
                    ConnectorPattern = shapeObject["connector_pattern"]?.ToString(),
                    ConnectorWeight = shapeObject["connector_weight"]?.Value<double>() ?? 0
                };

                Debug.WriteLine($"[CreateSingleShape] Processing shape:");
                Debug.WriteLine($"  Type: {shapeInfo.ShapeType}");
                Debug.WriteLine($"  ID: {shapeInfo.ShapeId}");
                Debug.WriteLine($"  Is Connector: {shapeInfo.IsConnector}");
                
                if (shapeInfo.IsConnector)
                {
                    // Ensure we have both source and target IDs for the connector
                    if (string.IsNullOrEmpty(shapeInfo.SourceShapeId) || string.IsNullOrEmpty(shapeInfo.TargetShapeId))
                    {
                        Debug.WriteLine("[CreateSingleShape] Error: Connector missing source_shape_id or target_shape_id");
                        return;
                    }

                    Debug.WriteLine($"  Source Shape ID: {shapeInfo.SourceShapeId}");
                    Debug.WriteLine($"  Target Shape ID: {shapeInfo.TargetShapeId}");
                    Debug.WriteLine($"  Connector Type: {shapeInfo.ConnectorType}");
                    
                    var connectorShape = libraryManager.ConnectShapes(
                        shapeInfo.SourceShapeId,
                        shapeInfo.TargetShapeId,
                        shapeInfo.ConnectorType ?? "Default"
                    );

                    if (connectorShape != null)
                    {
                        try
                        {
                            // Set connector properties only if they have valid values
                            try
                            {
                                Debug.WriteLine("[Connector] Setting properties:");
                                if (!string.IsNullOrEmpty(shapeInfo.ConnectorPattern))
                                {
                                    Debug.WriteLine($"[Connector] Setting line pattern to: {shapeInfo.ConnectorPattern}");
                                    try
                                    {
                                        connectorShape.get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject,
                                            (short)Visio.VisRowIndices.visRowLine,
                                            (short)Visio.VisCellIndices.visLinePattern).FormulaU = shapeInfo.ConnectorPattern;
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine($"[Connector] Failed to set line pattern: {ex.Message}");
                                    }
                                }
                                if (shapeInfo.ConnectorWeight > 0)
                                {
                                    Debug.WriteLine($"[Connector] Setting line weight to: {shapeInfo.ConnectorWeight}");
                                    try
                                    {
                                        connectorShape.get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject,
                                            (short)Visio.VisRowIndices.visRowLine,
                                            (short)Visio.VisCellIndices.visLineWeight).FormulaU = 
                                            shapeInfo.ConnectorWeight.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture) + " pt";
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine($"[Connector] Failed to set line weight: {ex.Message}");
                                    }
                                }
                                if (!string.IsNullOrEmpty(shapeInfo.ShapeColor))
                                {
                                    Debug.WriteLine($"[Connector] Setting color to: {shapeInfo.ShapeColor}");
                                    libraryManager.SetShapeColor(connectorShape, shapeInfo.ShapeColor);
                                }
                                if (!string.IsNullOrEmpty(shapeInfo.ShapeId))
                                {
                                    Debug.WriteLine($"[Connector] Setting name to: {shapeInfo.ShapeId}");
                                    connectorShape.NameU = shapeInfo.ShapeId;
                                }
                                Debug.WriteLine("[Connector] All properties set successfully");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[Connector] Error setting properties: {ex.Message}");
                                Debug.WriteLine($"[Connector] Stack trace: {ex.StackTrace}");
                                // Continue anyway since the connector is created and connected
                            }

                            currentCommandShapes.Add(shapeInfo);
                            Debug.WriteLine($"[CreateSingleShape] Successfully created connector between shapes {shapeInfo.SourceShapeId} and {shapeInfo.TargetShapeId}");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[CreateSingleShape] Warning: Some connector properties could not be set: {ex.Message}");
                            // Continue anyway since the connector is created and connected
                        }
                    }
                    else
                    {
                        Debug.WriteLine($"[CreateSingleShape] Failed to create connector between shapes {shapeInfo.SourceShapeId} and {shapeInfo.TargetShapeId}");
                    }
                    return;
                }

                // Add validation for required properties for non-connector shapes
                if (string.IsNullOrEmpty(shapeInfo.ShapeType))
                {
                    Debug.WriteLine("[CreateSingleShape] Error: shape_type is missing or empty");
                    return;
                }

                if (string.IsNullOrEmpty(shapeInfo.Category))
                {
                    Debug.WriteLine("[CreateSingleShape] Error: category is missing or empty");
                    return;
                }

                Debug.WriteLine($"[CreateSingleShape] Creating shape with properties:");
                Debug.WriteLine($"Category: {shapeInfo.Category}");
                Debug.WriteLine($"Shape Type: {shapeInfo.ShapeType}");
                Debug.WriteLine($"Position - X: {shapeInfo.PosX}%, Y: {shapeInfo.PosY}%");
                Debug.WriteLine($"Size - Width: {shapeInfo.Width}%, Height: {shapeInfo.Height}%");

                var shape = libraryManager.AddShapeToDocument(
                    shapeInfo.Category,
                    shapeInfo.ShapeType,
                    shapeInfo.PosX,
                    shapeInfo.PosY,
                    shapeInfo.Width,
                    shapeInfo.Height,
                    shapeInfo
                );

                if (shape != null)
                {
                    currentCommandShapes.Add(shapeInfo);
                    Debug.WriteLine($"[CreateSingleShape] Successfully created shape {shapeInfo.ShapeType}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[CreateSingleShape] Error: {ex.Message}");
                Debug.WriteLine($"[CreateSingleShape] Stack trace: {ex.StackTrace}");
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
                // Get page dimensions
                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                // Convert percentages to Visio units (inches)
                double visioX = (xPercent / 100.0) * pageWidth;
                double visioY = pageHeight - ((yPercent / 100.0) * pageHeight); // Adjust for Visio's coordinate system

                // Calculate initial size based on text length and font size
                double initialWidth = Math.Max(0.5, content.Length * fontSize * 0.08);
                double initialHeight = fontSize * 0.15;

                // Create a shape for text
                var textShape = activePage.DrawRectangle(
                    visioX - (initialWidth / 2),
                    visioY - (initialHeight / 2),
                    visioX + (initialWidth / 2),
                    visioY + (initialHeight / 2)
                );

                // Set text properties
                textShape.Text = content;

                // Set font size and color using CellsU instead of CharProps
                textShape.CellsU["Char.Size"].FormulaU = $"{fontSize} pt";
                textShape.CellsU["Char.Color"].FormulaU = $"RGB({ConvertColorToRGB(color)})";

                // Remove shape border and fill
                textShape.CellsU["LinePattern"].FormulaU = "0";
                textShape.CellsU["FillPattern"].FormulaU = "0";

                // Set text alignment
                textShape.CellsU["VerticalAlign"].FormulaU = "1"; // Middle
                textShape.CellsU["HAlign"].FormulaU = "1"; // Center

                // Allow text to resize shape
                textShape.CellsU["LockTextEdit"].FormulaU = "0";
                textShape.CellsU["LockWidth"].FormulaU = "0";
                textShape.CellsU["LockHeight"].FormulaU = "0";

                // Set text block properties
                textShape.CellsU["TxtWidth"].FormulaU = "Width*1";
                textShape.CellsU["TxtHeight"].FormulaU = "Height*1";
                textShape.CellsU["TxtPinX"].FormulaU = "Width*0.5";
                textShape.CellsU["TxtPinY"].FormulaU = "Height*0.5";

                Debug.WriteLine($"[ExecuteCreateTextBoxCommand] Created text with content: '{content}' at ({visioX}, {visioY})");

                // Track the created shape
                var shapeInfo = new VisioPlugin.ShapeInfo
                {
                    ShapeId = textShape.ID16.ToString(),
                    ShapeType = "Text",
                    Text = content
                };
                currentCommandShapes.Add(shapeInfo);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExecuteCreateTextBoxCommand] [Error] Failed to create text: {ex.Message}");
                throw; // Rethrow to ensure the error is properly reported
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