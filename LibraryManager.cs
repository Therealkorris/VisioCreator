using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioPlugin
{
    public class LibraryManager
    {
        private readonly Visio.Application visioApplication;
        private readonly Dictionary<string, ShapeCategory> categories;
        private HashSet<string> usedShapeIds = new HashSet<string>();  // Track used shape IDs

        public LibraryManager(Visio.Application visioApp)
        {
            visioApplication = visioApp ?? throw new ArgumentNullException(nameof(visioApp));
            categories = new Dictionary<string, ShapeCategory>();
            usedShapeIds = new HashSet<string>();
            LoadLibraries();
        }

        // Method to ensure unique shape IDs
        private string EnsureUniqueShapeId(string requestedId)
        {
            // If a specific ID is requested, always use it
            if (!string.IsNullOrEmpty(requestedId))
            {
                usedShapeIds.Add(requestedId);
                return requestedId;
            }

            // Only generate a new ID if none was provided
            int counter = 1;
            string newId = counter.ToString();
            while (usedShapeIds.Contains(newId))
            {
                counter++;
                newId = counter.ToString();
            }
            usedShapeIds.Add(newId);
            return newId;
        }

        // Method to register existing shape IDs (used when loading a document)
        private void RegisterExistingShapeId(string shapeId)
        {
            if (!string.IsNullOrEmpty(shapeId))
            {
                usedShapeIds.Add(shapeId);
            }
        }

        // Method to clear shape ID registry (used when closing/creating new documents)
        public void ClearShapeIdRegistry()
        {
            usedShapeIds.Clear();
        }

        public void LoadLibraries()
        {
            categories.Clear();
            BuildShapesCatalog();
        }

        public async Task SendShapesToN8n(string n8nWebhookUrl)
        {
            using (var client = new HttpClient())
            {
                var shapesCatalog = GetShapesCatalog();
                var jsonString = JsonConvert.SerializeObject(shapesCatalog);
                var content = new StringContent(jsonString, Encoding.UTF8, "application/json");

                try
                {
                    var response = await client.PostAsync(n8nWebhookUrl, content);
                    response.EnsureSuccessStatusCode();
                    Debug.WriteLine("[SendShapesToN8n] Shape catalog sent successfully.");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[SendShapesToN8n] Failed to send shape catalog: {ex.Message}");
                    // Consider adding more robust error handling here, e.g., retries, logging to a file, etc.
                }
            }
        }

        public JObject GetShapesCatalog()
        {
            var catalog = new JObject();
            foreach (var category in categories)
            {
                var shapesArray = new JArray();
                foreach (var shapeName in category.Value.GetShapeNames())
                {
                    shapesArray.Add(shapeName);
                }
                catalog[category.Key] = shapesArray;
            }
            return catalog;
        }

        private void BuildShapesCatalog()
        {
            if (visioApplication?.Documents == null)
            {
                Debug.WriteLine("[BuildShapesCatalog] Visio application or documents are null.");
                return;
            }

            foreach (Visio.Document stencilDoc in visioApplication.Documents)
            {
                if (stencilDoc.Type == Visio.VisDocumentTypes.visTypeStencil)
                {
                    // Use the full stencil name as the category
                    string category = stencilDoc.Name;
                    if (!categories.ContainsKey(category))
                    {
                        categories[category] = new ShapeCategory(category);
                    }

                    foreach (Visio.Master master in stencilDoc.Masters)
                    {
                        categories[category].AddShape(master.Name, master);
                    }
                }
            }
        }

        public IEnumerable<string> GetCategories()
        {
            return categories.Keys;
        }

        public IEnumerable<string> GetShapesInCategory(string categoryName)
        {
            if (categories.TryGetValue(categoryName, out ShapeCategory category))
            {
                return category.GetShapeNames();
            }
            return Enumerable.Empty<string>();
        }

        public Visio.Master GetShape(string categoryName, string shapeName)
        {
            try
            {
                // Clean up the shape name by removing any trailing ".number"
                string cleanShapeName = System.Text.RegularExpressions.Regex.Replace(shapeName, @"\.\d+$", "");
                Debug.WriteLine($"[GetShape] Looking for shape: {cleanShapeName} (original: {shapeName})");

                // First check if we already have the shape in our cache
                if (categories.TryGetValue(categoryName, out ShapeCategory category))
                {
                    var shape = category.GetShape(cleanShapeName);
                    if (shape != null)
                    {
                        return shape;
                    }
                }

                // If not found in cache, try to load from stencil
                var stencilDoc = visioApplication.Documents.OpenEx(categoryName, 
                    (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);
                
                if (stencilDoc != null)
                {
                    // Add category to cache if it doesn't exist
                    if (!categories.ContainsKey(categoryName))
                    {
                        categories[categoryName] = new ShapeCategory(categoryName);
                        foreach (Visio.Master master in stencilDoc.Masters)
                        {
                            categories[categoryName].AddShape(master.Name, master);
                        }
                    }

                    // Try exact match with cleaned name
                    foreach (Visio.Master master in stencilDoc.Masters)
                    {
                        if (master.Name == cleanShapeName)
                        {
                            return master;
                        }
                    }
                }

                Debug.WriteLine($"[GetShape] Shape not found: {cleanShapeName}");
                return null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[GetShape] Error: {ex.Message}");
                return null;
            }
        }

        public Visio.Master GetShapeByName(string shapeName)
        {
            foreach (var category in categories.Values)
            {
                var shape = category.GetShape(shapeName);
                if (shape != null)
                {
                    return shape;
                }
            }
            return null;
        }

        public Visio.Shape AddShapeToDocument(string category, string shapeName, double xPercent, double yPercent, double widthPercent, double heightPercent, ShapeInfo shapeInfo = null)
        {
            try
            {
                var activePage = visioApplication?.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[AddShapeToDocument] No active page found.");
                    return null;
                }

                // Get page dimensions
                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                // Convert percentages to Visio units (inches)
                double visioX = (xPercent / 100.0) * pageWidth;
                double visioY = (yPercent / 100.0) * pageHeight;
                double visioWidth = (widthPercent / 100.0) * pageWidth;
                double visioHeight = (heightPercent / 100.0) * pageHeight;

                // Get and drop regular shape
                var master = GetShape(category, shapeInfo?.ShapeType ?? shapeName);
                if (master == null)
                {
                    Debug.WriteLine($"[AddShapeToDocument] Master shape not found in {category}");
                    return null;
                }

                Visio.Shape shape = null;
                try
                {
                    shape = activePage.Drop(master, visioX, visioY);
                    if (shape == null)
                    {
                        Debug.WriteLine($"[AddShapeToDocument] Failed to drop shape {shapeName}");
                        return null;
                    }

                    try
                    {
                        Debug.WriteLine("[AddShapeToDocument] Setting Width");
                        shape.CellsU["Width"].ResultIU = visioWidth;
                        
                        Debug.WriteLine("[AddShapeToDocument] Setting Height");
                        shape.CellsU["Height"].ResultIU = visioHeight;
                        
                        Debug.WriteLine("[AddShapeToDocument] Setting PinX");
                        shape.CellsU["PinX"].ResultIU = visioX;
                        
                        Debug.WriteLine("[AddShapeToDocument] Setting PinY");
                        shape.CellsU["PinY"].ResultIU = visioY;

                        // Apply additional properties if provided
                        if (shapeInfo != null)
                        {
                            if (!string.IsNullOrEmpty(shapeInfo.ShapeColor))
                            {
                                Debug.WriteLine("[AddShapeToDocument] Setting Color");
                                SetShapeColor(shape, shapeInfo.ShapeColor);
                            }
                            if (!string.IsNullOrEmpty(shapeInfo.Text))
                            {
                                Debug.WriteLine("[AddShapeToDocument] Setting Text");
                                shape.Text = shapeInfo.Text;
                            }
                            if (shapeInfo.Angle != 0)
                            {
                                Debug.WriteLine("[AddShapeToDocument] Setting Angle");
                                shape.CellsU["Angle"].ResultIU = shapeInfo.Angle;
                            }

                            // Ensure unique shape ID and set it
                            try
                            {
                                // Use the original shape ID from the request
                                string shapeId = shapeInfo?.ShapeId;
                                if (string.IsNullOrEmpty(shapeId))
                                {
                                    // Only generate a new ID if none was provided
                                    shapeId = EnsureUniqueShapeId(null);
                                }

                                // First, set the shape's name
                                shape.Name = shapeId;
                                Debug.WriteLine($"[AddShapeToDocument] Set shape name to: {shapeId}");

                                // Add custom property section if it doesn't exist
                                if (shape.SectionExists[(short)Visio.VisSectionIndices.visSectionProp, 0] == 0)
                                {
                                    shape.AddSection((short)Visio.VisSectionIndices.visSectionProp);
                                }

                                // Find if ShapeId property already exists
                                short existingRow = -1;
                                var propSection = shape.Section[(short)Visio.VisSectionIndices.visSectionProp];
                                for (short propRow = 0; propRow < propSection.Count; propRow++)
                                {
                                    if (shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, propRow, (short)Visio.VisCellIndices.visCustPropsLabel].ResultStr[""] == "ShapeId")
                                    {
                                        existingRow = propRow;
                                        break;
                                    }
                                }

                                // Delete existing row if found
                                if (existingRow != -1)
                                {
                                    shape.DeleteRow((short)Visio.VisSectionIndices.visSectionProp, existingRow);
                                }

                                // Add new row for ShapeId
                                short newPropRow = shape.AddRow(
                                    (short)Visio.VisSectionIndices.visSectionProp,
                                    (short)Visio.VisRowIndices.visRowProp,
                                    (short)Visio.VisRowTags.visTagDefault
                                );

                                // Set the property name (label)
                                shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, newPropRow, (short)Visio.VisCellIndices.visCustPropsLabel].FormulaForceU = "\"ShapeId\"";

                                // Set the property value to the original ID
                                shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, newPropRow, (short)Visio.VisCellIndices.visCustPropsValue].FormulaForceU = $"\"{shapeId}\"";

                                // Set property type to 0 (string)
                                shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, newPropRow, (short)Visio.VisCellIndices.visCustPropsType].FormulaForceU = "0";

                                Debug.WriteLine($"[AddShapeToDocument] Successfully set shape ID property to: {shapeId}");

                                // Verify the property was set
                                string verifyValue = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, newPropRow, (short)Visio.VisCellIndices.visCustPropsValue].ResultStr[""];
                                Debug.WriteLine($"[AddShapeToDocument] Verified shape ID property value: {verifyValue}");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[AddShapeToDocument] Error setting shape ID property: {ex.Message}");
                            }
                        }

                        Debug.WriteLine($"[AddShapeToDocument] Successfully created shape {shapeName} from {category}");
                        return shape;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[AddShapeToDocument] Error setting properties: {ex.Message}");
                        return shape; // Return the shape anyway since it was created
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[AddShapeToDocument] Error dropping shape: {ex.Message}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AddShapeToDocument] Error: {ex.Message}");
                return null;
            }
        }

        // New and enhanced functions for greater Visio control:

        public Visio.Shape ConnectShapes(string sourceShapeId, string targetShapeId, string connectorType, string connectorShapeId = null)
        {
            try
            {
                var activePage = visioApplication?.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[Connector] No active page found.");
                    return null;
                }

                Debug.WriteLine($"[Connector] Looking for shapes with IDs {sourceShapeId} and {targetShapeId}");
                Debug.WriteLine($"[Connector] Total shapes on page: {activePage.Shapes.Count}");

                // Find shapes by their IDs
                Visio.Shape sourceShape = null;
                Visio.Shape targetShape = null;

                foreach (Visio.Shape shape in activePage.Shapes)
                {
                    try
                    {
                        // Get the ShapeId property if it exists
                        if (shape.SectionExists[(short)Visio.VisSectionIndices.visSectionProp, 0] != 0)
                        {
                            var propSection = shape.Section[(short)Visio.VisSectionIndices.visSectionProp];
                            for (short row = 0; row < propSection.Count; row++)
                            {
                                string propName = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, row, (short)Visio.VisCellIndices.visCustPropsLabel].ResultStr[""];
                                if (propName == "ShapeId")
                                {
                                    string propValue = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, row, (short)Visio.VisCellIndices.visCustPropsValue].ResultStr[""];
                                    Debug.WriteLine($"[Connector] Found shape with ShapeId property: {propValue}");
                                    
                                    if (propValue == sourceShapeId)
                                    {
                                        sourceShape = shape;
                                        Debug.WriteLine($"[Connector] Found source shape with ID {sourceShapeId}");
                                    }
                                    else if (propValue == targetShapeId)
                                    {
                                        targetShape = shape;
                                        Debug.WriteLine($"[Connector] Found target shape with ID {targetShapeId}");
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[Connector] Error checking shape properties: {ex.Message}");
                    }

                    if (sourceShape != null && targetShape != null)
                    {
                        break;
                    }
                }

                if (sourceShape == null || targetShape == null)
                {
                    Debug.WriteLine($"[Connector] Could not find shapes with IDs {sourceShapeId} and {targetShapeId}");
                    Debug.WriteLine($"[Connector] Source shape found: {sourceShape != null}");
                    Debug.WriteLine($"[Connector] Target shape found: {targetShape != null}");
                    return null;
                }

                Debug.WriteLine($"[Connector] Creating connector between {sourceShape.Name} and {targetShape.Name}");

                // Add a dynamic connector
                Debug.WriteLine("[Connector] Getting connector tool data");
                var connector = activePage.Application.ConnectorToolDataObject;
                Debug.WriteLine("[Connector] Dropping connector on page");
                var connectorShape = activePage.Drop(connector, 0, 0);

                if (connectorShape == null)
                {
                    Debug.WriteLine("[Connector] Failed to create connector shape");
                    return null;
                }

                Debug.WriteLine("[Connector] Gluing connector endpoints");

                try
                {
                    // Glue the connector's begin point to the first shape
                    Debug.WriteLine("[Connector] Gluing begin point");
                    connectorShape.CellsU["BeginX"].GlueTo(sourceShape.CellsU["PinX"]);
                    connectorShape.CellsU["BeginY"].GlueTo(sourceShape.CellsU["PinY"]);

                    // Glue the connector's end point to the second shape
                    Debug.WriteLine("[Connector] Gluing end point");
                    connectorShape.CellsU["EndX"].GlueTo(targetShape.CellsU["PinX"]);
                    connectorShape.CellsU["EndY"].GlueTo(targetShape.CellsU["PinY"]);

                    // Set the connector type if needed (e.g., straight, curved)
                    if (!string.IsNullOrEmpty(connectorType))
                    {
                        Debug.WriteLine($"[Connector] Setting connector type to: {connectorType}");
                        switch (connectorType.ToLower())
                        {
                            case "curved":
                                connectorShape.CellsU["ShapeRouteStyle"].FormulaU = "2";
                                break;
                            case "straight":
                                connectorShape.CellsU["ShapeRouteStyle"].FormulaU = "1";
                                break;
                            case "dynamic":
                                connectorShape.CellsU["ShapeRouteStyle"].FormulaU = "16";
                                break;
                            default:
                                connectorShape.CellsU["ShapeRouteStyle"].FormulaU = "1"; // Default to straight
                                break;
                        }
                    }

                    // Add custom property section if it doesn't exist
                    if (connectorShape.SectionExists[(short)Visio.VisSectionIndices.visSectionProp, 0] == 0)
                    {
                        connectorShape.AddSection((short)Visio.VisSectionIndices.visSectionProp);
                    }

                    // Add ShapeId property to the connector
                    short newPropRow = connectorShape.AddRow(
                        (short)Visio.VisSectionIndices.visSectionProp,
                        (short)Visio.VisRowIndices.visRowProp,
                        (short)Visio.VisRowTags.visTagDefault
                    );

                    // Set the property name (label)
                    connectorShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, newPropRow, (short)Visio.VisCellIndices.visCustPropsLabel].FormulaForceU = "\"ShapeId\"";

                    // Set the property value to the provided connector shape ID
                    connectorShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, newPropRow, (short)Visio.VisCellIndices.visCustPropsValue].FormulaForceU = $"\"{connectorShapeId}\"";
                    connectorShape.Name = connectorShapeId; // Also set the shape's name to match

                    // Set property type to 0 (string)
                    connectorShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, newPropRow, (short)Visio.VisCellIndices.visCustPropsType].FormulaForceU = "0";

                    // Add source and target shape IDs as custom properties
                    short sourceRow = connectorShape.AddRow(
                        (short)Visio.VisSectionIndices.visSectionProp,
                        (short)Visio.VisRowIndices.visRowProp,
                        (short)Visio.VisRowTags.visTagDefault
                    );
                    connectorShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, sourceRow, (short)Visio.VisCellIndices.visCustPropsLabel].FormulaForceU = "\"SourceShapeId\"";
                    connectorShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, sourceRow, (short)Visio.VisCellIndices.visCustPropsValue].FormulaForceU = $"\"{sourceShapeId}\"";
                    connectorShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, sourceRow, (short)Visio.VisCellIndices.visCustPropsType].FormulaForceU = "0";

                    short targetRow = connectorShape.AddRow(
                        (short)Visio.VisSectionIndices.visSectionProp,
                        (short)Visio.VisRowIndices.visRowProp,
                        (short)Visio.VisRowTags.visTagDefault
                    );
                    connectorShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, targetRow, (short)Visio.VisCellIndices.visCustPropsLabel].FormulaForceU = "\"TargetShapeId\"";
                    connectorShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, targetRow, (short)Visio.VisCellIndices.visCustPropsValue].FormulaForceU = $"\"{targetShapeId}\"";
                    connectorShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, targetRow, (short)Visio.VisCellIndices.visCustPropsType].FormulaForceU = "0";

                    Debug.WriteLine($"[Connector] Successfully connected shapes with IDs {sourceShapeId} and {targetShapeId}");
                    Debug.WriteLine($"[Connector] Set connector ShapeId to: {connectorShapeId}");
                    return connectorShape;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[Connector] Error gluing connector: {ex.Message}");
                    Debug.WriteLine($"[Connector] Stack trace: {ex.StackTrace}");
                    
                    // Try to clean up the failed connector
                    try
                    {
                        if (connectorShape != null)
                        {
                            connectorShape.Delete();
                            Debug.WriteLine("[Connector] Cleaned up failed connector");
                        }
                    }
                    catch { }
                    
                    return null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connector] Error connecting shapes: {ex.Message}");
                Debug.WriteLine($"[Connector] Stack trace: {ex.StackTrace}");
                return null;
            }
        }

        public Visio.Shape AddTextToShape(string shapeName, string text)
        {
            try
            {
                var shape = visioApplication.ActivePage.Shapes.ItemU[shapeName];
                shape.Text = text;
                Debug.WriteLine($"Added text '{text}' to shape: {shapeName}");
                return shape;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error adding text to shape: {ex.Message}");
                return null;
            }
        }

        public Visio.Shape SetShapeStyle(string shapeName, string lineStyle, string fillPattern)
        {
            try
            {
                var shape = visioApplication.ActivePage.Shapes.ItemU[shapeName];
                if (!string.IsNullOrEmpty(lineStyle))
                {
                    shape.CellsU["LinePattern"].FormulaU = lineStyle;
                }
                if (!string.IsNullOrEmpty(fillPattern))
                {
                    shape.CellsU["FillPattern"].FormulaU = fillPattern;
                }
                Debug.WriteLine($"Set style for shape: {shapeName} (Line: {lineStyle}, Fill: {fillPattern})");
                return shape;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error setting shape style: {ex.Message}");
                return null;
            }
        }

        public Visio.Shape GroupShapes(string[] shapeNames)
        {
            try
            {
                var activePage = visioApplication.ActivePage;
                var selection = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty);
                foreach (var shapeName in shapeNames)
                {
                    selection.Select(activePage.Shapes.ItemU[shapeName], (short)Visio.VisSelectArgs.visSelect);
                }
                var groupedShape = selection.Group();
                Debug.WriteLine($"Grouped shapes: {string.Join(", ", shapeNames)} into {groupedShape.Name}");
                return groupedShape;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error grouping shapes: {ex.Message}");
                return null;
            }
        }

        public List<Visio.Shape> UngroupShapes(string shapeName)
        {
            try
            {
                var shape = visioApplication.ActivePage.Shapes.ItemU[shapeName];
                var ungroupedShapes = new List<Visio.Shape>();
                
                // Create a selection and select the shape to ungroup
                var selection = visioApplication.ActiveWindow.Selection;
                selection.DeselectAll();
                selection.Select(shape, (short)Visio.VisSelectArgs.visSelect);
                
                // Ungroup and collect the resulting shapes
                selection.Ungroup();
                foreach (Visio.Shape ungroupedShape in visioApplication.ActiveWindow.Selection)
                {
                    ungroupedShapes.Add(ungroupedShape);
                }
                
                Debug.WriteLine($"Ungrouped shape: {shapeName}");
                return ungroupedShapes;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error ungrouping shape: {ex.Message}");
                return new List<Visio.Shape>();
            }
        }

        public List<Visio.Shape> AlignShapes(string[] shapeNames, string alignmentType)
        {
            try
            {
                var activePage = visioApplication.ActivePage;
                var selection = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty);
                var shapes = new List<Visio.Shape>();
                foreach (var shapeName in shapeNames)
                {
                    var shape = activePage.Shapes.ItemU[shapeName];
                    selection.Select(shape, (short)Visio.VisSelectArgs.visSelect);
                    shapes.Add(shape);
                }

                switch (alignmentType.ToLower())
                {
                    case "left":
                        selection.Align(Visio.VisHorizontalAlignTypes.visHorzAlignLeft, Visio.VisVerticalAlignTypes.visVertAlignNone, true);
                        break;
                    case "center":
                        selection.Align(Visio.VisHorizontalAlignTypes.visHorzAlignCenter, Visio.VisVerticalAlignTypes.visVertAlignNone, true);
                        break;
                    case "right":
                        selection.Align(Visio.VisHorizontalAlignTypes.visHorzAlignRight, Visio.VisVerticalAlignTypes.visVertAlignNone, true);
                        break;
                    case "top":
                        selection.Align(Visio.VisHorizontalAlignTypes.visHorzAlignNone, Visio.VisVerticalAlignTypes.visVertAlignTop, true);
                        break;
                    case "middle":
                        selection.Align(Visio.VisHorizontalAlignTypes.visHorzAlignNone, Visio.VisVerticalAlignTypes.visVertAlignMiddle, true);
                        break;
                    case "bottom":
                        selection.Align(Visio.VisHorizontalAlignTypes.visHorzAlignNone, Visio.VisVerticalAlignTypes.visVertAlignBottom, true);
                        break;
                }

                Debug.WriteLine($"Aligned shapes: {string.Join(", ", shapeNames)} to {alignmentType}");
                return shapes;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error aligning shapes: {ex.Message}");
                return new List<Visio.Shape>();
            }
        }

        public List<Visio.Shape> DistributeShapes(string[] shapeNames, string distributionType)
        {
            try
            {
                var activePage = visioApplication.ActivePage;
                var selection = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty);
                var shapes = new List<Visio.Shape>();
                foreach (var shapeName in shapeNames)
                {
                    var shape = activePage.Shapes.ItemU[shapeName];
                    selection.Select(shape, (short)Visio.VisSelectArgs.visSelect);
                    shapes.Add(shape);
                }

                switch (distributionType.ToLower())
                {
                    case "horizontal":
                        selection.Distribute(Visio.VisDistributeTypes.visDistHorzSpace, true);
                        break;
                    case "vertical":
                        selection.Distribute(Visio.VisDistributeTypes.visDistVertSpace, true);
                        break;
                }

                Debug.WriteLine($"Distributed shapes: {string.Join(", ", shapeNames)} {distributionType}");
                return shapes;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error distributing shapes: {ex.Message}");
                return new List<Visio.Shape>();
            }
        }

        public string GetShapeProperties(string shapeName)
        {
            try
            {
                var shape = visioApplication.ActivePage.Shapes.ItemU[shapeName];
                var properties = new
                {
                    Name = shape.Name,
                    Type = shape.Master?.Name ?? "No Master",
                    Position = new { X = shape.CellsU["PinX"].ResultIU, Y = shape.CellsU["PinY"].ResultIU },
                    Size = new { Width = shape.CellsU["Width"].ResultIU, Height = shape.CellsU["Height"].ResultIU },
                    Color = shape.CellsU["FillForegnd"].FormulaU,
                    Text = shape.Text
                };
                return JsonConvert.SerializeObject(properties);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting shape properties: {ex.Message}");
                return JsonConvert.SerializeObject(new { error = ex.Message });
            }
        }

        public string GetPageSize()
        {
            try
            {
                var activePage = visioApplication.ActivePage;
                var pageSize = new
                {
                    Width = activePage.PageSheet.CellsU["PageWidth"].ResultIU,
                    Height = activePage.PageSheet.CellsU["PageHeight"].ResultIU
                };
                return JsonConvert.SerializeObject(pageSize);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting page size: {ex.Message}");
                return JsonConvert.SerializeObject(new { error = ex.Message });
            }
        }

        // Helper method to scale a percentage (0-100) to page dimension
        private double ScaleToPageDimension(double percent, double dimension)
        {
            return (percent / 100.0) * dimension;
        }

        public void SetShapeColor(Visio.Shape shape, string colorHex)
        {
            try
            {
                if (shape == null || string.IsNullOrEmpty(colorHex))
                {
                    return;
                }

                var color = System.Drawing.ColorTranslator.FromHtml(colorHex);
                string rgbValue = $"{color.R},{color.G},{color.B}";

                shape.CellsU["FillForegnd"].FormulaU = $"RGB({rgbValue})";
                shape.CellsU["LineColor"].FormulaU = $"RGB({rgbValue})";
                shape.CellsU["FillPattern"].FormulaU = "1";

                Debug.WriteLine($"Set color for shape '{shape.Name}' to {colorHex}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error setting color for shape '{shape.Name}': {ex.Message}");
            }
        }

        public string GetShapeColor(Visio.Shape shape)
        {
            try
            {
                string rgbColor = shape.CellsU["FillForegnd"].ResultStr[""];
                if (string.IsNullOrEmpty(rgbColor)) return "";

                // Parse RGB values
                var rgbMatch = System.Text.RegularExpressions.Regex.Match(rgbColor, @"RGB\((\d+);\s*(\d+);\s*(\d+)\)");
                if (rgbMatch.Success)
                {
                    int r = int.Parse(rgbMatch.Groups[1].Value);
                    int g = int.Parse(rgbMatch.Groups[2].Value);
                    int b = int.Parse(rgbMatch.Groups[3].Value);
                    return $"#{r:X2}{g:X2}{b:X2}";
                }
                return rgbColor;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting color for shape '{shape.Name}': {ex.Message}");
                return "";
            }
        }

        public List<VisioPlugin.ShapeInfo> ListAllShapes()
        {
            var shapes = new List<VisioPlugin.ShapeInfo>();
            var activePage = visioApplication?.ActivePage;
            if (activePage == null) return shapes;

            // Clear existing shape ID registry before reading shapes
            ClearShapeIdRegistry();

            // Set page dimensions for percentage calculations
            double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
            double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;
            VisioPlugin.ShapeInfo.SetPageDimensions(pageWidth, pageHeight);

            // First pass: Collect all shapes and their basic information
            foreach (Visio.Shape shape in activePage.Shapes)
            {
                string shapeId = null;
                string shapeType = null;
                string sourceShapeId = null;
                string targetShapeId = null;

                // Get the custom ShapeId property and original shape type
                if (shape.SectionExists[(short)Visio.VisSectionIndices.visSectionProp, 0] != 0)
                {
                    var propSection = shape.Section[(short)Visio.VisSectionIndices.visSectionProp];
                    for (short row = 0; row < propSection.Count; row++)
                    {
                        string propName = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, row, (short)Visio.VisCellIndices.visCustPropsLabel].ResultStr[""];
                        string propValue = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, row, (short)Visio.VisCellIndices.visCustPropsValue].ResultStr[""];
                        
                        switch (propName)
                        {
                            case "ShapeId":
                                shapeId = propValue;
                                RegisterExistingShapeId(shapeId); // Register the shape ID
                                break;
                            case "SourceShapeId":
                                sourceShapeId = propValue;
                                break;
                            case "TargetShapeId":
                                targetShapeId = propValue;
                                break;
                        }
                    }
                }

                // If no custom ShapeId found, generate a new unique one
                if (string.IsNullOrEmpty(shapeId))
                {
                    shapeId = EnsureUniqueShapeId(null);
                    Debug.WriteLine($"[ListAllShapes] Generated new shape ID for shape without one: {shapeId}");
                }

                // Get the original shape type from the master
                if (shape.Master != null)
                {
                    // Get the original master name without any Visio-added numbers
                    shapeType = System.Text.RegularExpressions.Regex.Replace(shape.Master.Name, @"\.\d+$", "");
                }
                else
                {
                    // For connectors or other special shapes
                    shapeType = shape.Name;
                    if (shape.CellExists["BeginX", 0] != 0 && shape.CellExists["EndX", 0] != 0)
                    {
                        shapeType = "Dynamic connector";
                    }
                }

                Debug.WriteLine($"[ListAllShapes] Processing shape: ID={shapeId}, Type={shapeType}, Master={(shape.Master != null ? shape.Master.Name : "null")}");

                var shapeInfo = new VisioPlugin.ShapeInfo
                {
                    ShapeId = shapeId,
                    ShapeType = shapeType,
                    ShapeColor = GetShapeColor(shape),
                    Text = shape.Text,
                    PosX = shape.CellsU["PinX"].ResultIU,
                    PosY = shape.CellsU["PinY"].ResultIU,
                    Width = shape.CellsU["Width"].ResultIU,
                    Height = shape.CellsU["Height"].ResultIU,
                    Angle = shape.CellsU["Angle"].ResultIU,
                    ZOrder = shape.Index,
                    Category = GetOriginalStencilName(shape)
                };

                // Check if it's a connector
                bool isConnector = shape.CellExists["BeginX", 0] != 0 && shape.CellExists["EndX", 0] != 0;
                shapeInfo.IsConnector = isConnector;

                if (isConnector)
                {
                    shapeInfo.BeginX = shape.CellsU["BeginX"].ResultIU;
                    shapeInfo.BeginY = shape.CellsU["BeginY"].ResultIU;
                    shapeInfo.EndX = shape.CellsU["EndX"].ResultIU;
                    shapeInfo.EndY = shape.CellsU["EndY"].ResultIU;
                    
                    // First try to get source and target IDs from custom properties
                    if (!string.IsNullOrEmpty(sourceShapeId) && !string.IsNullOrEmpty(targetShapeId))
                    {
                        shapeInfo.SourceShapeId = sourceShapeId;
                        shapeInfo.TargetShapeId = targetShapeId;
                        Debug.WriteLine($"Found source/target IDs from properties - Source: {sourceShapeId}, Target: {targetShapeId}");
                    }
                    // If not found in properties, try to get them from connections
                    else
                    {
                        try
                        {
                            foreach (Visio.Connect connect in shape.Connects)
                            {
                                string cellName = connect.FromCell.Name.ToLower();
                                Visio.Shape connectedShape = connect.ToSheet;

                                // Get the ShapeId property of the connected shape
                                string connectedShapeId = null;
                                if (connectedShape.SectionExists[(short)Visio.VisSectionIndices.visSectionProp, 0] != 0)
                                {
                                    var propSection = connectedShape.Section[(short)Visio.VisSectionIndices.visSectionProp];
                                    for (short row = 0; row < propSection.Count; row++)
                                    {
                                        if (connectedShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, row, (short)Visio.VisCellIndices.visCustPropsLabel].ResultStr[""] == "ShapeId")
                                        {
                                            connectedShapeId = connectedShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionProp, row, (short)Visio.VisCellIndices.visCustPropsValue].ResultStr[""];
                                            break;
                                        }
                                    }
                                }

                                if (!string.IsNullOrEmpty(connectedShapeId))
                                {
                                    if (cellName.Contains("beginx") || cellName.Contains("beginy"))
                                    {
                                        shapeInfo.SourceShapeId = connectedShapeId;
                                        Debug.WriteLine($"Found source shape from connection: {connectedShapeId}");
                                    }
                                    else if (cellName.Contains("endx") || cellName.Contains("endy"))
                                    {
                                        shapeInfo.TargetShapeId = connectedShapeId;
                                        Debug.WriteLine($"Found target shape from connection: {connectedShapeId}");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error getting connector endpoints: {ex.Message}");
                        }
                    }
                    
                    // Get connector type and styling
                    try
                    {
                        string routeStyle = shape.CellsU["ShapeRouteStyle"].ResultStr[""];
                        // Map the route style to meaningful values
                        shapeInfo.ConnectorType = routeStyle switch
                        {
                            "0" => "Default",
                            "1" => "Straight",
                            "2" => "Curved",
                            "3" => "Right Angle",
                            "4" => "Curved Right Angle",
                            _ => routeStyle
                        };

                        // Get line pattern and weight
                        shapeInfo.ConnectorPattern = shape.CellsU["LinePattern"].ResultStr[""];
                        shapeInfo.ConnectorWeight = shape.CellsU["LineWeight"].ResultIU;
                        shapeInfo.ConnectorRounding = shape.CellsU["Rounding"].ResultStr[""];

                        // Get routing points only if the shape has a geometry section
                        if (shape.SectionExists[(short)Visio.VisSectionIndices.visSectionFirstComponent, 0] != 0)
                        {
                            var geomSection = shape.Section[(short)Visio.VisSectionIndices.visSectionFirstComponent];
                            for (short row = 0; row < geomSection.Count; row++)
                            {
                                try
                                {
                                    // Get X and Y coordinates directly without trying to access row names
                                    double x = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, row, (short)Visio.VisCellIndices.visX].ResultIU;
                                    double y = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, row, (short)Visio.VisCellIndices.visY].ResultIU;
                                    
                                    shapeInfo.RoutingPoints.Add(new VisioPlugin.Point(x, y));
                                    Debug.WriteLine($"Added routing point: ({x}, {y})");

                                    // Try to get control points if they exist
                                    try
                                    {
                                        double control1X = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, row, (short)Visio.VisCellIndices.visControl1X].ResultIU;
                                        double control1Y = shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionFirstComponent, row, (short)Visio.VisCellIndices.visControl1Y].ResultIU;
                                        shapeInfo.ControlPoints.Add(new VisioPlugin.ControlPoint(control1X, control1Y, "Control1"));
                                    }
                                    catch
                                    {
                                        // Control points don't exist for this vertex - this is normal
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"Error processing geometry point at row {row}: {ex.Message}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error getting connector styling: {ex.Message}");
                        shapeInfo.ConnectorType = "Default";
                    }
                }

                // Get custom properties
                try
                {
                    if (shape.SectionExists[(short)Visio.VisSectionIndices.visSectionProp, 0] != 0)
                    {
                        var propSection = shape.Section[(short)Visio.VisSectionIndices.visSectionProp];
                        var rowCount = propSection.Count;

                        for (short row = 0; row < rowCount; row++)
                        {
                            string propName = shape.CellsSRC[
                                (short)Visio.VisSectionIndices.visSectionProp,
                                row,
                                (short)Visio.VisCellIndices.visCustPropsLabel
                            ].ResultStr[""];

                            string propValue = shape.CellsSRC[
                                (short)Visio.VisSectionIndices.visSectionProp,
                                row,
                                (short)Visio.VisCellIndices.visCustPropsValue
                            ].ResultStr[""];
                            
                            if (!string.IsNullOrEmpty(propName))
                            {
                                shapeInfo.CustomProperties[propName] = propValue;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error getting custom properties for shape {shape.Name}: {ex.Message}");
                }

                shapes.Add(shapeInfo);
                Debug.WriteLine($"Added shape: {shapeInfo}");
            }

            return shapes;
        }

        // Helper method to get the original stencil name
        private string GetOriginalStencilName(Visio.Shape shape)
        {
            try
            {
                if (shape.Master == null) return "Dynamic";

                // Try to get the original stencil name from the master's document
                var masterDoc = shape.Master.Document;
                
                // Check if this is a document stencil or a regular stencil
                if (masterDoc.Type == Visio.VisDocumentTypes.visTypeStencil)
                {
                    // Use the full document name including extension
                    return masterDoc.Name;
                }
                else
                {
                    // For document stencils or other cases, try to find the original category
                    foreach (var category in categories)
                    {
                        if (category.Value.GetShape(shape.Master.Name) != null)
                        {
                            return category.Key;
                        }
                    }
                    return "Document Stencil";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting stencil name for shape: {ex.Message}");
                return "Unknown";
            }
        }

        public void ProcessSingleCommand(Dictionary<string, object> command)
        {
            try
            {
                string commandType = command["CommandType"] as string;
                if (commandType == "CreateShape")
                {
                    var shapes = command["Parameters"] as Dictionary<string, object>;
                    var shapesList = shapes?["shapes"] as List<VisioPlugin.ShapeInfo>;
                    if (shapesList != null)
                    {
                        // Split shapes into non-connectors and connectors
                        var nonConnectors = new List<VisioPlugin.ShapeInfo>();
                        var connectors = new List<VisioPlugin.ShapeInfo>();

                        foreach (var shape in shapesList)
                        {
                            if (shape.ShapeType.ToLower().Contains("connector"))
                            {
                                connectors.Add(shape);
                            }
                            else
                            {
                                nonConnectors.Add(shape);
                            }
                        }

                        // Process non-connectors first
                        foreach (var shape in nonConnectors)
                        {
                            Debug.WriteLine($"[CreateSingleShape] Processing regular shape:\n  Type: {shape.ShapeType}\n  ID: {shape.ShapeId}");
                            CreateSingleShape(shape);
                        }

                        // Then process connectors
                        foreach (var shape in connectors)
                        {
                            Debug.WriteLine($"[CreateSingleShape] Processing connector:\n  Type: {shape.ShapeType}\n  ID: {shape.ShapeId}\n  Source: {shape.SourceShapeId}\n  Target: {shape.TargetShapeId}");
                            CreateSingleShape(shape);
                        }
                    }
                }
                // ... rest of the existing code ...
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessSingleCommand] Error: {ex.Message}");
            }
        }

        private Visio.Shape CreateSingleShape(VisioPlugin.ShapeInfo shapeInfo)
        {
            try
            {
                if (shapeInfo.IsConnector)
                {
                    Debug.WriteLine($"[CreateSingleShape] Creating connector with ID {shapeInfo.ShapeId} between shapes {shapeInfo.SourceShapeId} and {shapeInfo.TargetShapeId}");
                    var connector = ConnectShapes(shapeInfo.SourceShapeId, shapeInfo.TargetShapeId, shapeInfo.ConnectorType, shapeInfo.ShapeId);
                    if (connector != null)
                    {
                        // Apply connector properties
                        if (!string.IsNullOrEmpty(shapeInfo.ShapeColor))
                        {
                            SetShapeColor(connector, shapeInfo.ShapeColor);
                        }
                        if (!string.IsNullOrEmpty(shapeInfo.ConnectorPattern))
                        {
                            connector.CellsU["LinePattern"].Formula = shapeInfo.ConnectorPattern;
                        }
                        if (shapeInfo.ConnectorWeight > 0)
                        {
                            connector.CellsU["LineWeight"].ResultIU = shapeInfo.ConnectorWeight;
                        }
                        
                        Debug.WriteLine($"[CreateSingleShape] Successfully created connector");
                        return connector;
                    }
                    else
                    {
                        Debug.WriteLine($"[CreateSingleShape] Failed to create connector between shapes {shapeInfo.SourceShapeId} and {shapeInfo.TargetShapeId}");
                        return null;
                    }
                }
                else
                {
                    Debug.WriteLine($"[CreateSingleShape] Creating shape with properties:\nCategory: {shapeInfo.Category}\nShape Type: {shapeInfo.ShapeType}\nPosition - X: {shapeInfo.PosX}%, Y: {shapeInfo.PosY}%\nSize - Width: {shapeInfo.Width}%, Height: {shapeInfo.Height}%");
                    return AddShapeToDocument(shapeInfo.Category, shapeInfo.ShapeType, shapeInfo.PosX, shapeInfo.PosY, shapeInfo.Width, shapeInfo.Height, shapeInfo);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[CreateSingleShape] Error: {ex.Message}");
                return null;
            }
        }
    }

    public class ShapeCategory
    {
        public string Name { get; }
        private readonly Dictionary<string, Visio.Master> shapes;

        public ShapeCategory(string name)
        {
            Name = name;
            shapes = new Dictionary<string, Visio.Master>();
        }

        public void AddShape(string name, Visio.Master master)
        {
            shapes[name] = master;
        }

        public IEnumerable<string> GetShapeNames()
        {
            return shapes.Keys;
        }

        public Visio.Master GetShape(string name)
        {
            return shapes.TryGetValue(name, out Visio.Master master) ? master : null;
        }
    }

    public class Position
    {
        public double X { get; set; }
        public double Y { get; set; }
    }
}