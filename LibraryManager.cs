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
            if (categories.TryGetValue(categoryName, out ShapeCategory category))
            {
                return category.GetShape(shapeName);
            }
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

        public Visio.Shape AddShapeToDocument(string categoryName, string shapeName, double xPercent, double yPercent, double widthPercent, double heightPercent)
        {
            try
            {
                Debug.WriteLine($"[AddShapeToDocument] Adding shape: {shapeName} from category: {categoryName} at ({xPercent}%, {yPercent}%) with size ({widthPercent}%, {heightPercent}%)");

                var activePage = visioApplication?.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[AddShapeToDocument] [Error] No active page found in Visio application.");
                    return null;
                }

                // Get the page dimensions in Visio internal units (inches)
                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                var master = GetShape(categoryName, shapeName);
                if (master == null)
                {
                    Debug.WriteLine($"[AddShapeToDocument] [Error] Shape '{shapeName}' not found in category '{categoryName}'.");
                    return null;
                }

                // Convert percentages to Visio units
                // Note: Visio uses inches internally
                double shapeWidth = (widthPercent / 100.0) * pageWidth;
                double shapeHeight = (heightPercent / 100.0) * pageHeight;

                // Calculate position in Visio units
                // Adjust for Visio's coordinate system (origin at bottom-left)
                double xPos = (xPercent / 100.0) * pageWidth;
                double yPos = pageHeight - ((yPercent / 100.0) * pageHeight);

                Debug.WriteLine($"[AddShapeToDocument] Page dimensions (inches) - Width: {pageWidth}, Height: {pageHeight}");
                Debug.WriteLine($"[AddShapeToDocument] Position (inches) - X: {xPos}, Y: {yPos}");
                Debug.WriteLine($"[AddShapeToDocument] Size (inches) - Width: {shapeWidth}, Height: {shapeHeight}");

                // Create the shape at the calculated position
                var shape = activePage.Drop(master, xPos, yPos);

                // Set the shape's size
                shape.Cells["Width"].ResultIU = shapeWidth;
                shape.Cells["Height"].ResultIU = shapeHeight;

                // Ensure the shape is centered on the target position
                shape.Cells["PinX"].ResultIU = xPos;
                shape.Cells["PinY"].ResultIU = yPos;

                // Lock aspect ratio for consistent shape appearance
                shape.Cells["LockAspect"].Formula = "1";

                Debug.WriteLine($"[AddShapeToDocument] Final position (inches) - PinX: {shape.Cells["PinX"].ResultIU}, PinY: {shape.Cells["PinY"].ResultIU}");

                return shape;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AddShapeToDocument] [Error] Error adding shape '{shapeName}' from category '{categoryName}': {ex.Message}");
                Debug.WriteLine($"Stack Trace: {ex.StackTrace}");
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
                    ZOrder = shape.Index
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