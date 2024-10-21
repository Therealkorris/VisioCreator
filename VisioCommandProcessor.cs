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
            // Add more commands if needed in the future
        }

        // The core command processor method
        public void ProcessCommand(string jsonCommand)
        {
            try
            {
                Debug.WriteLine($"Received Command: {jsonCommand}"); // Log the received command

                // Parse the JSON command
                JObject commandObject = JObject.Parse(jsonCommand);
                string commandName = commandObject["command"]?.ToString();

                if (string.IsNullOrEmpty(commandName))
                    throw new Exception("Command name is missing.");

                // Check if the command is registered
                if (commandRegistry.ContainsKey(commandName))
                {
                    Debug.WriteLine($"Executing Command: {commandName}"); // Log command execution
                    commandRegistry[commandName](commandObject["parameters"]);
                }
                else
                {
                    throw new Exception($"Command '{commandName}' is not recognized.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error processing command: {ex.Message}");
                throw; // Re-throw the exception to be handled upstream
            }
        }

        // New method to process JSON commands
        public void ProcessJsonCommand(string jsonCommand)
        {
            try
            {
                Debug.WriteLine($"Received JSON Command: {jsonCommand}"); // Log the received command

                // Parse the JSON command
                JObject commandObject = JObject.Parse(jsonCommand);
                string commandName = commandObject["command"]?.ToString();

                if (string.IsNullOrEmpty(commandName))
                    throw new Exception("Command name is missing.");

                // Check if the command is registered
                if (commandRegistry.ContainsKey(commandName))
                {
                    Debug.WriteLine($"Executing JSON Command: {commandName}"); // Log command execution
                    commandRegistry[commandName](commandObject["parameters"]);
                }
                else
                {
                    throw new Exception($"Command '{commandName}' is not recognized.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error processing JSON command: {ex.Message}");
                throw; // Re-throw the exception to be handled upstream
            }
        }

        // Command methods
        private void CreateShape(JToken parameters)
        {
            try
            {
                string shapeType = parameters["shapeType"]?.ToString();
                float xPercent = parameters["position"]?["x"]?.Value<float>() ?? 50;
                float yPercent = parameters["position"]?["y"]?.Value<float>() ?? 50;
                float widthPercent = parameters["size"]?["width"]?.Value<float>() ?? 10;
                float heightPercent = parameters["size"]?["height"]?.Value<float>() ?? 10;
                string color = parameters["color"]?.ToString();

                // Get the current category from your AI command or the current selection in the app
                string categoryName = Globals.ThisAddIn.CurrentCategory;

                if (string.IsNullOrEmpty(categoryName))
                {
                    Debug.WriteLine("[Error] No category specified. Cannot add shape.");
                    return;
                }

                // Convert percentage to Visio coordinates
                var activePage = visioApp.ActivePage;
                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                double visioX = (xPercent / 100.0) * pageWidth;
                double visioY = ((100 - yPercent) / 100.0) * pageHeight;
                double visioWidth = (widthPercent / 100.0) * pageWidth;
                double visioHeight = (heightPercent / 100.0) * pageHeight;

                // Add the shape using the category (stencil) and shape type
                libraryManager.AddShapeToDocument(categoryName, shapeType, visioX, visioY, visioWidth, visioHeight);

                // Get the last added shape to set properties (like color)
                Visio.Shape addedShape = activePage.Shapes.Cast<Visio.Shape>().LastOrDefault();

                // Set color if provided
                if (addedShape != null && !string.IsNullOrEmpty(color))
                {
                    libraryManager.SetShapeColor(addedShape, color);
                }

                Debug.WriteLine($"Shape '{shapeType}' created at ({visioX}, {visioY}) with size ({visioWidth}, {visioHeight}) and color {color}.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error creating shape: {ex.Message}");
                throw; // Re-throw the exception to be handled upstream
            }
        }
    }
}
