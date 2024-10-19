﻿using System;
using System.Reflection;
using Newtonsoft.Json.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic; // <-- Important for Dictionary

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
            // Mapping command names to their respective method handlers
            commandRegistry.Add("CreateShape", CreateShape);
            commandRegistry.Add("DeleteShape", DeleteShape);
            commandRegistry.Add("ConnectShapes", ConnectShapes);

            // Add more commands here as needed
        }

        // The core command processor method
        public void ProcessCommand(string jsonCommand)
        {
            try
            {
                Debug.WriteLine($"Received Command: {jsonCommand}"); // Log the received command

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


        // Command methods
        private void CreateShape(JToken parameters)
        {
            string shapeType = parameters["shapeType"]?.ToString();
            float xPercent = parameters["position"]?["x"]?.Value<float>() ?? 50;
            float yPercent = parameters["position"]?["y"]?.Value<float>() ?? 50;
            float widthPercent = parameters["size"]?["width"]?.Value<float>() ?? 10;
            float heightPercent = parameters["size"]?["height"]?.Value<float>() ?? 10;
            string color = parameters["color"]?.ToString();

            // Convert percentage to Visio coordinates
            var activePage = visioApp.ActivePage;
            double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
            double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

            double visioX = (xPercent / 100.0) * pageWidth;
            double visioY = ((100 - yPercent) / 100.0) * pageHeight;
            double visioWidth = (widthPercent / 100.0) * pageWidth;
            double visioHeight = (heightPercent / 100.0) * pageHeight;

            // No need to specify category; LibraryManager will find the shape
            libraryManager.AddShapeToDocument(null, shapeType, visioX, visioY, visioWidth, visioHeight);

            // Get the last added shape to set properties
            Visio.Shape addedShape = activePage.Shapes.Cast<Visio.Shape>().LastOrDefault();

            // Set color if provided
            if (addedShape != null && !string.IsNullOrEmpty(color))
            {
                libraryManager.SetShapeColor(addedShape, color);
            }

            Debug.WriteLine($"Shape '{shapeType}' created at ({visioX}, {visioY}) with size ({visioWidth}, {visioHeight}) and color {color}.");
        }

        private void DeleteShape(JToken parameters)
        {
            string shapeName = parameters["shapeName"]?.ToString();

            if (string.IsNullOrEmpty(shapeName))
            {
                Debug.WriteLine("Shape name is missing for DeleteShape command.");
                return;
            }

            var activePage = visioApp.ActivePage;
            var shape = activePage.Shapes.Cast<Visio.Shape>()
                .FirstOrDefault(s => string.Equals(s.Name, shapeName, StringComparison.OrdinalIgnoreCase));

            if (shape != null)
            {
                shape.Delete();
                Debug.WriteLine($"Shape '{shapeName}' deleted.");
            }
            else
            {
                Debug.WriteLine($"Shape '{shapeName}' not found on the active page.");
            }
        }

        private void ConnectShapes(JToken parameters)
        {
            string fromShapeName = parameters["fromShape"]?.ToString();
            string toShapeName = parameters["toShape"]?.ToString();
            string connectorType = parameters["connectorType"]?.ToString() ?? "Dynamic Connector";

            if (string.IsNullOrEmpty(fromShapeName) || string.IsNullOrEmpty(toShapeName))
            {
                Debug.WriteLine("FromShape or ToShape is missing for ConnectShapes command.");
                return;
            }

            var activePage = visioApp.ActivePage;
            var fromShape = activePage.Shapes.Cast<Visio.Shape>()
                .FirstOrDefault(s => string.Equals(s.Name, fromShapeName, StringComparison.OrdinalIgnoreCase));
            var toShape = activePage.Shapes.Cast<Visio.Shape>()
                .FirstOrDefault(s => string.Equals(s.Name, toShapeName, StringComparison.OrdinalIgnoreCase));

            if (fromShape == null || toShape == null)
            {
                Debug.WriteLine($"One or both shapes '{fromShapeName}', '{toShapeName}' not found.");
                return;
            }

            // Get the connector master
            Visio.Master connectorMaster = visioApp.ConnectorToolDataObject as Visio.Master;
            if (connectorMaster == null)
            {
                Debug.WriteLine("Connector master not found.");
                return;
            }

            // Drop the connector
            Visio.Shape connector = activePage.Drop(connectorMaster, 0, 0);

            // Connect the shapes
            connector.CellsU["BeginX"].GlueTo(fromShape.CellsU["PinX"]);
            connector.CellsU["EndX"].GlueTo(toShape.CellsU["PinX"]);

            Debug.WriteLine($"Shapes '{fromShapeName}' and '{toShapeName}' connected with '{connectorType}'.");
        }
    }
}