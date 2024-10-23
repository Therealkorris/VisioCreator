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
                string shapeName = parameters["shapeName"]?.ToString();

                Debug.WriteLine($"[DeleteShape] ShapeName: {shapeName}");

                // Get the current active Visio page
                var activePage = visioApp.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[DeleteShape] [Error] No active page found in Visio.");
                    return;
                }

                // Find the shape by name and delete it
                Visio.Shape shapeToDelete = activePage.Shapes.Cast<Visio.Shape>().FirstOrDefault(s => s.Name == shapeName);
                if (shapeToDelete != null)
                {
                    shapeToDelete.Delete();
                    Debug.WriteLine($"[DeleteShape] Shape '{shapeName}' deleted successfully.");
                }
                else
                {
                    Debug.WriteLine($"[DeleteShape] [Error] Shape '{shapeName}' not found.");
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
                string shapeName = parameters["shapeName"]?.ToString();
                float xPercent = parameters["position"]?["x"]?.Value<float>() ?? 50;  // Percentage x-coordinate
                float yPercent = parameters["position"]?["y"]?.Value<float>() ?? 50;  // Percentage y-coordinate

                Debug.WriteLine($"[MoveShape] ShapeName: {shapeName}, X: {xPercent}%, Y: {yPercent}%");

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

                // Find the shape by name and move it
                Visio.Shape shapeToMove = activePage.Shapes.Cast<Visio.Shape>().FirstOrDefault(s => s.Name == shapeName);
                if (shapeToMove != null)
                {
                    shapeToMove.CellsU["PinX"].ResultIU = visioX;
                    shapeToMove.CellsU["PinY"].ResultIU = visioY;
                    Debug.WriteLine($"[MoveShape] Shape '{shapeName}' moved successfully.");
                }
                else
                {
                    Debug.WriteLine($"[MoveShape] [Error] Shape '{shapeName}' not found.");
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
    }
}
