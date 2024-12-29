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

        public LibraryManager(Visio.Application visioApp)
        {
            visioApplication = visioApp ?? throw new ArgumentNullException(nameof(visioApp));
            categories = new Dictionary<string, ShapeCategory>();
            LoadLibraries();
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
                        Debug.WriteLine($"Added shape '{master.Name}' from stencil '{category}'");
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
            Debug.WriteLine($"[GetShape] Looking for shape '{shapeName}' in category '{categoryName}'");
            
            // First try to get from loaded stencils
            if (categories.TryGetValue(categoryName, out ShapeCategory category))
            {
                var shape = category.GetShape(shapeName);
                if (shape != null)
                {
                    Debug.WriteLine($"[GetShape] Found shape in loaded stencils");
                    return shape;
                }
            }

            // If not found, try to force-load the stencil
            try
            {
                Debug.WriteLine($"[GetShape] Attempting to force-load stencil: {categoryName}");
                var stencilDoc = visioApplication.Documents.OpenEx(categoryName, 
                    (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);
                
                if (stencilDoc != null)
                {
                    // Find the master by name
                    foreach (Visio.Master master in stencilDoc.Masters)
                    {
                        if (master.Name == shapeName)
                        {
                            Debug.WriteLine($"[GetShape] Found shape in force-loaded stencil");
                            return master;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[GetShape] Error force-loading stencil: {ex.Message}");
            }

            Debug.WriteLine($"[GetShape] Shape not found");
            return null;
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

                // Get the master shape
                var master = GetShape(category, shapeName);
                if (master == null)
                {
                    // Try to find the shape in any loaded stencil
                    foreach (Visio.Document doc in visioApplication.Documents)
                    {
                        if (doc.Type == Visio.VisDocumentTypes.visTypeStencil)
                        {
                            foreach (Visio.Master m in doc.Masters)
                            {
                                if (m.Name == shapeName)
                                {
                                    master = m;
                                    break;
                                }
                            }
                            if (master != null) break;
                        }
                    }

                    if (master == null)
                    {
                        Debug.WriteLine($"[AddShapeToDocument] Master shape not found: {shapeName} in {category}");
                        return null;
                    }
                }

                // Drop the shape at the specified position
                var shape = activePage.Drop(master, visioX, visioY);
                if (shape == null)
                {
                    Debug.WriteLine("[AddShapeToDocument] Failed to drop shape.");
                    return null;
                }

                // Set basic properties
                shape.CellsU["Width"].ResultIU = visioWidth;
                shape.CellsU["Height"].ResultIU = visioHeight;
                shape.CellsU["PinX"].ResultIU = visioX;
                shape.CellsU["PinY"].ResultIU = visioY;

                // Apply additional properties if ShapeInfo is provided
                if (shapeInfo != null)
                {
                    // Set color
                    if (!string.IsNullOrEmpty(shapeInfo.ShapeColor))
                    {
                        SetShapeColor(shape, shapeInfo.ShapeColor);
                    }

                    // Set text
                    if (!string.IsNullOrEmpty(shapeInfo.Text))
                    {
                        shape.Text = shapeInfo.Text;
                    }

                    // Set angle
                    if (shapeInfo.Angle != 0)
                    {
                        shape.CellsU["Angle"].ResultIU = shapeInfo.Angle;
                    }

                    // Set z-order
                    if (shapeInfo.ZOrder != 0)
                    {
                        shape.CellsU["ZOrderIndex"].Formula = shapeInfo.ZOrder.ToString();
                    }

                    // Handle connector properties if it's a connector
                    if (shapeInfo.IsConnector)
                    {
                        // Set connector pattern
                        if (!string.IsNullOrEmpty(shapeInfo.ConnectorPattern))
                        {
                            shape.CellsU["LinePattern"].Formula = shapeInfo.ConnectorPattern;
                        }

                        // Set connector weight
                        if (shapeInfo.ConnectorWeight > 0)
                        {
                            shape.CellsU["LineWeight"].ResultIU = shapeInfo.ConnectorWeight;
                        }

                        // Set connector rounding
                        if (!string.IsNullOrEmpty(shapeInfo.ConnectorRounding))
                        {
                            shape.CellsU["Rounding"].Formula = shapeInfo.ConnectorRounding;
                        }

                        // Set begin and end points
                        if (shapeInfo.BeginX != 0 || shapeInfo.BeginY != 0)
                        {
                            shape.CellsU["BeginX"].ResultIU = (shapeInfo.BeginX / 100.0) * pageWidth;
                            shape.CellsU["BeginY"].ResultIU = (shapeInfo.BeginY / 100.0) * pageHeight;
                        }
                        if (shapeInfo.EndX != 0 || shapeInfo.EndY != 0)
                        {
                            shape.CellsU["EndX"].ResultIU = (shapeInfo.EndX / 100.0) * pageWidth;
                            shape.CellsU["EndY"].ResultIU = (shapeInfo.EndY / 100.0) * pageHeight;
                        }

                        // Connect to source and target shapes if specified
                        if (!string.IsNullOrEmpty(shapeInfo.SourceShapeId) && !string.IsNullOrEmpty(shapeInfo.TargetShapeId))
                        {
                            var sourceShape = activePage.Shapes.ItemFromID[int.Parse(shapeInfo.SourceShapeId)];
                            var targetShape = activePage.Shapes.ItemFromID[int.Parse(shapeInfo.TargetShapeId)];
                            if (sourceShape != null && targetShape != null)
                            {
                                shape.CellsU["BeginX"].GlueTo(sourceShape.CellsU["PinX"]);
                                shape.CellsU["BeginY"].GlueTo(sourceShape.CellsU["PinY"]);
                                shape.CellsU["EndX"].GlueTo(targetShape.CellsU["PinX"]);
                                shape.CellsU["EndY"].GlueTo(targetShape.CellsU["PinY"]);
                            }
                        }
                    }

                    // Set custom properties
                    foreach (var prop in shapeInfo.CustomProperties)
                    {
                        try
                        {
                            shape.AddNamedRow(
                                (short)Visio.VisSectionIndices.visSectionProp, 
                                prop.Key, 
                                (short)Visio.VisRowTags.visTagDefault
                            );
                            shape.CellsSRC[
                                (short)Visio.VisSectionIndices.visSectionProp, 
                                (short)(shape.RowCount[(short)Visio.VisSectionIndices.visSectionProp] - 1), 
                                (short)Visio.VisCellIndices.visCustPropsValue
                            ].FormulaU = $"\"{prop.Value}\"";
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[AddShapeToDocument] Error setting custom property {prop.Key}: {ex.Message}");
                        }
                    }
                }

                return shape;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AddShapeToDocument] Error: {ex.Message}");
                return null;
            }
        }

        // New and enhanced functions for greater Visio control:

        public Visio.Shape ConnectShapes(string shape1Name, string shape2Name, string connectorType)
        {
            try
            {
                var activePage = visioApplication.ActivePage;
                var shape1 = activePage.Shapes.ItemU[shape1Name];
                var shape2 = activePage.Shapes.ItemU[shape2Name];

                // Add a dynamic connector
                var connector = activePage.Application.ConnectorToolDataObject;
                var connectorShape = activePage.Drop(connector, 0, 0);

                // Glue the connector's begin point to the first shape
                connectorShape.CellsU["BeginX"].GlueTo(shape1.CellsU["PinX"]);

                // Glue the connector's end point to the second shape
                connectorShape.CellsU["EndX"].GlueTo(shape2.CellsU["PinX"]);

                // Set the connector type if needed (e.g., straight, curved)
                if (!string.IsNullOrEmpty(connectorType))
                {
                    // You might need to adjust this based on how connector types are represented in Visio
                    if (connectorType.Equals("curved", StringComparison.OrdinalIgnoreCase))
                    {
                        connectorShape.CellsU["ShapeRouteStyle"].FormulaU = "2"; // Example value for curved connectors
                    }
                    else
                    {
                        connectorShape.CellsU["ShapeRouteStyle"].FormulaU = "1"; // Example value for straight connectors
                    }
                }

                Debug.WriteLine($"Connected shapes: {shape1Name} and {shape2Name} with connector type: {connectorType}");
                return connectorShape;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error connecting shapes: {ex.Message}");
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

            // Set page dimensions for percentage calculations
            double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
            double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;
            VisioPlugin.ShapeInfo.SetPageDimensions(pageWidth, pageHeight);

            // First pass: Collect all shapes and their basic information
            foreach (Visio.Shape shape in activePage.Shapes)
            {
                var shapeInfo = new VisioPlugin.ShapeInfo
                {
                    ShapeId = shape.ID16.ToString(),
                    ShapeType = shape.Name,
                    ShapeColor = GetShapeColor(shape),
                    Text = shape.Text,
                    PosX = shape.CellsU["PinX"].ResultIU,
                    PosY = shape.CellsU["PinY"].ResultIU,
                    Width = shape.CellsU["Width"].ResultIU,
                    Height = shape.CellsU["Height"].ResultIU,
                    Angle = shape.CellsU["Angle"].ResultIU,
                    ZOrder = shape.Index,
                    // Get the original stencil name by checking the master's document
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

                    // Get connected shapes
                    try
                    {
                        // Examine each connection to determine source and target
                        foreach (Visio.Connect connect in shape.Connects)
                        {
                            // The FromCell will tell us if this is a begin or end point
                            string cellName = connect.FromCell.Name.ToLower();
                            
                            // The ToSheet is the shape we're connected to
                            Visio.Shape connectedShape = connect.ToSheet;
                            
                            if (cellName.Contains("beginx") || cellName.Contains("beginy"))
                            {
                                shapeInfo.SourceShapeId = connectedShape.ID16.ToString();
                                Debug.WriteLine($"Found source shape: {shapeInfo.SourceShapeId} via {cellName}");
                            }
                            else if (cellName.Contains("endx") || cellName.Contains("endy"))
                            {
                                shapeInfo.TargetShapeId = connectedShape.ID16.ToString();
                                Debug.WriteLine($"Found target shape: {shapeInfo.TargetShapeId} via {cellName}");
                            }
                        }

                        Debug.WriteLine($"Connector {shape.ID16} - Source: {shapeInfo.SourceShapeId}, Target: {shapeInfo.TargetShapeId}");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error getting connector endpoints: {ex.Message}");
                    }
                }

                // Get custom properties only if the shape has a properties section
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