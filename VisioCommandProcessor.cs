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
            commandRegistry.Add("UpdateShapeColor", UpdateShapeColor);
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
                float x = parameters["position"]?["x"]?.Value<float>() ?? 50;  // Absolute x-coordinate
                float y = parameters["position"]?["y"]?.Value<float>() ?? 50;  // Absolute y-coordinate
                float width = parameters["size"]?["width"]?.Value<float>() ?? 10;  // Absolute width
                float height = parameters["size"]?["height"]?.Value<float>() ?? 10;  // Absolute height
                string color = parameters["color"]?.ToString();

                Debug.WriteLine($"[CreateShape] ShapeType: {shapeType}, X: {x}, Y: {y}, Width: {width}, Height: {height}, Color: {color}");

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

                // Use provided coordinates directly as Visio coordinates
                double visioX = x;
                double visioY = y;

                // Use the provided width and height directly
                double visioWidth = width;
                double visioHeight = height;

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

        // Process the UpdateShapeColor command from AI
        private void UpdateShapeColor(JToken parameters)
        {
            try
            {
                Debug.WriteLine($"[UpdateShapeColor] Received Parameters: {parameters.ToString()}");

                // Extract parameters from the AI response
                string shapeName = parameters["shapeName"]?.ToString();
                string color = parameters["color"]?.ToString();

                Debug.WriteLine($"[UpdateShapeColor] ShapeName: {shapeName}, Color: {color}");

                // Get the current active Visio page
                var activePage = visioApp.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[UpdateShapeColor] [Error] No active page found in Visio.");
                    return;
                }

                // Find the shape by name
                Visio.Shape shape = activePage.Shapes.Cast<Visio.Shape>().FirstOrDefault(s => s.Name == shapeName);

                if (shape != null)
                {
                    // Apply color if it's specified in the AI command
                    if (!string.IsNullOrEmpty(color))
                    {
                        libraryManager.SetShapeColor(shape, color);
                        Debug.WriteLine($"[UpdateShapeColor] Applied color '{color}' to shape '{shape.Name}'.");
                    }
                }
                else
                {
                    Debug.WriteLine($"[UpdateShapeColor] [Error] Shape '{shapeName}' not found.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[UpdateShapeColor] [Error] Error updating shape color: {ex.Message}");
                throw;
            }
        }
    }
}
