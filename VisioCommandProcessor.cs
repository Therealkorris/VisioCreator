using System;
using System.Linq;
using Newtonsoft.Json.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using System.Diagnostics;
using System.Collections.Generic;

namespace VisioPlugin
{
    public class VisioCommandProcessor
    {
        private readonly Visio.Application visioApp;
        private readonly LibraryManager libraryManager;

        // Command registry to store command names and their corresponding handlers
        private Dictionary<string, Action<JToken>> commandRegistry;

        public VisioCommandProcessor(Visio.Application visioApp, LibraryManager libraryManager)
        {
            this.visioApp = visioApp;
            this.libraryManager = libraryManager;

            // Initialize the command registry
            commandRegistry = new Dictionary<string, Action<JToken>>(StringComparer.OrdinalIgnoreCase);

            // Register commands and their handlers
            RegisterCommands();
        }

        // Register all available commands dynamically
        private void RegisterCommands()
        {
            commandRegistry.Add("CreateShape", CreateShape);
            commandRegistry.Add("DeleteShape", DeleteShape);
            commandRegistry.Add("MoveShape", MoveShape);
            commandRegistry.Add("ResizeShape", ResizeShape);
            commandRegistry.Add("ConnectShapes", ConnectShapes);
            commandRegistry.Add("CreateText", CreateText);
            commandRegistry.Add("RetrieveAllShapes", RetrieveAllShapes);
        }

        // The core command processor method
        public void ProcessCommand(string jsonCommand)
        {
            try
            {
                Debug.WriteLine($"[ProcessCommand] Received Command: {jsonCommand}");

                // Parse the JSON command
                JObject commandObject = JObject.Parse(jsonCommand);
                string commandName = commandObject["command"]?.ToString();

                if (string.IsNullOrEmpty(commandName))
                    throw new Exception("[ProcessCommand] Command name is missing.");

                // Check if the command is registered
                if (commandRegistry.ContainsKey(commandName))
                {
                    Debug.WriteLine($"[ProcessCommand] Executing Command: {commandName}");
                    commandRegistry[commandName](commandObject["parameters"]);
                }
                else
                {
                    throw new Exception($"[ProcessCommand] Command '{commandName}' is not recognized.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessCommand] Error processing command: {ex.Message}");
                throw;
            }
        }

        // Process the CreateShape command from AI
        private void CreateShape(JToken parameters)
        {
            try
            {
                Debug.WriteLine($"[CreateShape] Received Parameters: {parameters.ToString()}");

                // Extract parameters from the AI response
                string shapeType = parameters["shapeType"]?.ToString();
                float xPercent = parameters["position"]?["x"]?.Value<float>() ?? 50;  // Percentage x-coordinate
                float yPercent = parameters["position"]?["y"]?.Value<float>() ?? 50;  // Percentage y-coordinate
                float widthPercent = parameters["size"]?["width"]?.Value<float>() ?? 10;  // Percentage width
                float heightPercent = parameters["size"]?["height"]?.Value<float>() ?? 10;  // Percentage height
                string color = parameters["color"]?.ToString();

                Debug.WriteLine($"[CreateShape] ShapeType: {shapeType}, X: {xPercent}%, Y: {yPercent}%, Width: {widthPercent}%, Height: {heightPercent}%, Color: {color}");

                // Get the current category (stencil) to be used to create the shape
                string categoryName = Globals.ThisAddIn.CurrentCategory; // Use CurrentCategory

                if (string.IsNullOrEmpty(categoryName))
                {
                    Debug.WriteLine("[CreateShape] [Error] No category specified. Cannot add shape.");
                    return;
                }

                Debug.WriteLine($"[CreateShape] Using Category: {categoryName}");

                // Get the current active Visio page
                var activePage = visioApp.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[CreateShape] [Error] No active page found in Visio.");
                    return;
                }

                // Retrieve the canvas dimensions
                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                // Convert percentage coordinates to absolute coordinates
                double visioX = (xPercent / 100.0) * pageWidth;
                double visioY = (1 - (yPercent / 100.0)) * pageHeight;

                // Ensure the coordinates fit within the canvas
                visioX = Math.Max(0, Math.Min(visioX, pageWidth));
                visioY = Math.Max(0, Math.Min(visioY, pageHeight));

                // Convert percentage size to absolute size
                double visioWidth = (widthPercent / 100.0) * pageWidth;
                double visioHeight = (heightPercent / 100.0) * pageHeight;

                Debug.WriteLine($"[CreateShape] Calculated Coordinates - X: {visioX}, Y: {visioY}, Width: {visioWidth}, Height: {visioHeight}");

                // Add the shape using the category (stencil) and shape type
                Debug.WriteLine($"[CreateShape] Attempting to add shape '{shapeType}' from category '{categoryName}'");
                libraryManager.AddShapeToDocument(categoryName, shapeType, visioX, visioY, visioWidth, visioHeight);

                // Get the last added shape to set additional properties (like color)
                Visio.Shape addedShape = activePage.Shapes.Cast<Visio.Shape>().LastOrDefault();

                if (addedShape != null)
                {
                    Debug.WriteLine($"[CreateShape] Shape '{addedShape.Name}' added successfully.");

                    // Apply color if it's specified in the AI command
                    if (!string.IsNullOrEmpty(color))
                    {
                        libraryManager.SetShapeColor(addedShape, color);
                        Debug.WriteLine($"[CreateShape] Applied color '{color}' to shape '{addedShape.Name}'.");
                    }
                }
                else
                {
                    Debug.WriteLine("[CreateShape] [Error] Shape was not added successfully.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[CreateShape] [Error] Error creating shape: {ex.Message}");
                throw;
            }
        }

        // Process the DeleteShape command from AI
        private void DeleteShape(JToken parameters)
        {
            try
            {
                Debug.WriteLine($"[DeleteShape] Received Parameters: {parameters.ToString()}");

                // Extract parameters from the AI response
                string color = parameters["color"]?.ToString();
                string shapeType = parameters["shapeType"]?.ToString();

                Debug.WriteLine($"[DeleteShape] Color: {color}, ShapeType: {shapeType}");

                // Get the current active Visio page
                var activePage = visioApp.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[DeleteShape] [Error] No active page found in Visio.");
                    return;
                }

                // Find the shape by color and shapeType and delete it
                Visio.Shape shapeToDelete = activePage.Shapes.Cast<Visio.Shape>().FirstOrDefault(s => s.CellsU["FillForegnd"].FormulaU.Contains(color) && s.Name.Contains(shapeType));
                if (shapeToDelete != null)
                {
                    shapeToDelete.Delete();
                    Debug.WriteLine($"[DeleteShape] Shape with Color '{color}' and ShapeType '{shapeType}' deleted successfully.");
                }
                else
                {
                    Debug.WriteLine($"[DeleteShape] [Error] Shape with Color '{color}' and ShapeType '{shapeType}' not found.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DeleteShape] [Error] Error deleting shape: {ex.Message}");
                throw;
            }
        }

        // Process the MoveShape command from AI
        private void MoveShape(JToken parameters)
        {
            try
            {
                Debug.WriteLine($"[MoveShape] Received Parameters: {parameters.ToString()}");

                // Extract parameters from the AI response
                string color = parameters["color"]?.ToString();
                string shapeType = parameters["shapeType"]?.ToString();
                float xPercent = parameters["position"]?["x"]?.Value<float>() ?? 50;  // Percentage x-coordinate
                float yPercent = parameters["position"]?["y"]?.Value<float>() ?? 50;  // Percentage y-coordinate

                Debug.WriteLine($"[MoveShape] Color: {color}, ShapeType: {shapeType}, X: {xPercent}%, Y: {yPercent}%");

                // Get the current active Visio page
                var activePage = visioApp.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[MoveShape] [Error] No active page found in Visio.");
                    return;
                }

                // Retrieve the canvas dimensions
                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                // Convert percentage coordinates to absolute coordinates
                double visioX = (xPercent / 100.0) * pageWidth;
                double visioY = (1 - (yPercent / 100.0)) * pageHeight;

                // Ensure the coordinates fit within the canvas
                visioX = Math.Max(0, Math.Min(visioX, pageWidth));
                visioY = Math.Max(0, Math.Min(visioY, pageHeight));

                Debug.WriteLine($"[MoveShape] Calculated Coordinates - X: {visioX}, Y: {visioY}");

                // Find the shape by color and shapeType and move it
                Visio.Shape shapeToMove = activePage.Shapes.Cast<Visio.Shape>().FirstOrDefault(s => s.CellsU["FillForegnd"].FormulaU.Contains(color) && s.Name.Contains(shapeType));
                if (shapeToMove != null)
                {
                    shapeToMove.CellsU["PinX"].ResultIU = visioX;
                    shapeToMove.CellsU["PinY"].ResultIU = visioY;
                    Debug.WriteLine($"[MoveShape] Shape with Color '{color}' and ShapeType '{shapeType}' moved successfully.");
                }
                else
                {
                    Debug.WriteLine($"[MoveShape] [Error] Shape with Color '{color}' and ShapeType '{shapeType}' not found.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[MoveShape] [Error] Error moving shape: {ex.Message}");
                throw;
            }
        }

        // Process the ResizeShape command from AI
        private void ResizeShape(JToken parameters)
        {
            try
            {
                Debug.WriteLine($"[ResizeShape] Received Parameters: {parameters.ToString()}");

                // Extract parameters from the AI response
                string shapeName = parameters["shapeName"]?.ToString();
                float widthPercent = parameters["size"]?["width"]?.Value<float>() ?? 10;  // Percentage width
                float heightPercent = parameters["size"]?["height"]?.Value<float>() ?? 10;  // Percentage height

                Debug.WriteLine($"[ResizeShape] ShapeName: {shapeName}, Width: {widthPercent}%, Height: {heightPercent}%");

                // Get the current active Visio page
                var activePage = visioApp.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[ResizeShape] [Error] No active page found in Visio.");
                    return;
                }

                // Retrieve the canvas dimensions
                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                // Convert percentage size to absolute size
                double visioWidth = (widthPercent / 100.0) * pageWidth;
                double visioHeight = (heightPercent / 100.0) * pageHeight;

                Debug.WriteLine($"[ResizeShape] Calculated Size - Width: {visioWidth}, Height: {visioHeight}");

                // Find the shape by name and resize it
                Visio.Shape shapeToResize = activePage.Shapes.Cast<Visio.Shape>().FirstOrDefault(s => s.Name == shapeName);
                if (shapeToResize != null)
                {
                    shapeToResize.CellsU["Width"].ResultIU = visioWidth;
                    shapeToResize.CellsU["Height"].ResultIU = visioHeight;
                    Debug.WriteLine($"[ResizeShape] Shape '{shapeName}' resized successfully.");
                }
                else
                {
                    Debug.WriteLine($"[ResizeShape] [Error] Shape '{shapeName}' not found.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ResizeShape] [Error] Error resizing shape: {ex.Message}");
                throw;
            }
        }

        // Process the ConnectShapes command from AI
        private void ConnectShapes(JToken parameters)
        {
            try
            {
                Debug.WriteLine($"[ConnectShapes] Received Parameters: {parameters.ToString()}");

                // Extract parameters from the AI response
                string shapeName1 = parameters["shapeName1"]?.ToString();
                string shapeName2 = parameters["shapeName2"]?.ToString();

                Debug.WriteLine($"[ConnectShapes] ShapeName1: {shapeName1}, ShapeName2: {shapeName2}");

                // Get the current active Visio page
                var activePage = visioApp.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[ConnectShapes] [Error] No active page found in Visio.");
                    return;
                }

                // Find the shapes by name
                Visio.Shape shape1 = activePage.Shapes.Cast<Visio.Shape>().FirstOrDefault(s => s.Name == shapeName1);
                Visio.Shape shape2 = activePage.Shapes.Cast<Visio.Shape>().FirstOrDefault(s => s.Name == shapeName2);

                if (shape1 != null && shape2 != null)
                {
                    // Create a dynamic connector
                    Visio.Shape connector = activePage.Drop(visioApp.ConnectorToolDataObject, 0, 0);

                    // Connect the shapes
                    connector.CellsU["BeginX"].GlueTo(shape1.CellsU["PinX"]);
                    connector.CellsU["EndX"].GlueTo(shape2.CellsU["PinX"]);

                    Debug.WriteLine($"[ConnectShapes] Shapes '{shapeName1}' and '{shapeName2}' connected successfully.");
                }
                else
                {
                    Debug.WriteLine($"[ConnectShapes] [Error] One or both shapes not found. Shape1: {shapeName1}, Shape2: {shapeName2}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ConnectShapes] [Error] Error connecting shapes: {ex.Message}");
                throw;
            }
        }

        // Process the CreateText command from AI
        private void CreateText(JToken parameters)
        {
            try
            {
                Debug.WriteLine($"[CreateText] Received Parameters: {parameters.ToString()}");

                // Extract parameters from the AI response
                string textContent = parameters["textContent"]?.ToString();
                float xPercent = parameters["position"]?["x"]?.Value<float>() ?? 50;  // Percentage x-coordinate
                float yPercent = parameters["position"]?["y"]?.Value<float>() ?? 50;  // Percentage y-coordinate
                float fontSize = parameters["fontSize"]?.Value<float>() ?? 12;  // Font size
                string color = parameters["color"]?.ToString();

                Debug.WriteLine($"[CreateText] TextContent: {textContent}, X: {xPercent}%, Y: {yPercent}%, FontSize: {fontSize}, Color: {color}");

                // Get the current active Visio page
                var activePage = visioApp.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[CreateText] [Error] No active page found in Visio.");
                    return;
                }

                // Retrieve the canvas dimensions
                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                // Convert percentage coordinates to absolute coordinates
                double visioX = (xPercent / 100.0) * pageWidth;
                double visioY = (1 - (yPercent / 100.0)) * pageHeight;

                // Ensure the coordinates fit within the canvas
                visioX = Math.Max(0, Math.Min(visioX, pageWidth));
                visioY = Math.Max(0, Math.Min(visioY, pageHeight));

                Debug.WriteLine($"[CreateText] Calculated Coordinates - X: {visioX}, Y: {visioY}");

                // Add the text to the Visio page
                Visio.Shape textShape = activePage.DrawRectangle(visioX, visioY, visioX + 1, visioY + 1);
                textShape.Text = textContent;
                textShape.CellsU["Char.Size"].ResultIU = fontSize;

                // Apply color if it's specified in the AI command
                if (!string.IsNullOrEmpty(color))
                {
                    textShape.CellsU["Char.Color"].FormulaU = $"RGB({color})";
                    Debug.WriteLine($"[CreateText] Applied color '{color}' to text.");
                }

                Debug.WriteLine($"[CreateText] Text '{textContent}' added successfully.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[CreateText] [Error] Error adding text: {ex.Message}");
                throw;
            }
        }

        // Retrieve all shapes in the active Visio page
        public void RetrieveAllShapes(JToken parameters)
        {
            try
            {
                Debug.WriteLine($"[RetrieveAllShapes] Received Parameters: {parameters.ToString()}");

                // Get the current active Visio page
                var activePage = visioApp.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[RetrieveAllShapes] [Error] No active page found in Visio.");
                    return;
                }

                // Iterate through all shapes in the active page and collect their properties
                var shapes = activePage.Shapes.Cast<Visio.Shape>().Select(shape => new
                {
                    Name = shape.Name,
                    Type = shape.Master.Name,
                    Position = new
                    {
                        X = shape.CellsU["PinX"].ResultIU,
                        Y = shape.CellsU["PinY"].ResultIU
                    },
                    Color = shape.CellsU["FillForegnd"].FormulaU
                }).ToList();

                // Log the retrieved shapes
                Debug.WriteLine($"[RetrieveAllShapes] Retrieved {shapes.Count} shapes.");
                foreach (var shape in shapes)
                {
                    Debug.WriteLine($"[RetrieveAllShapes] Shape - Name: {shape.Name}, Type: {shape.Type}, Position: ({shape.Position.X}, {shape.Position.Y}), Color: {shape.Color}");
                }

                // Send the shapes data back to the AI
                // Note: Implement the logic to send the shapes data back to the AI as needed
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[RetrieveAllShapes] [Error] Error retrieving shapes: {ex.Message}");
                throw;
            }
        }
    }
}
