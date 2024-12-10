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

        public void AddShapeToDocument(string categoryName, string shapeName, double xPercent, double yPercent, double widthPercent, double heightPercent)
        {
            try
            {
                Debug.WriteLine($"[AddShapeToDocument] Adding shape: {shapeName} from category: {categoryName} at ({xPercent}%, {yPercent}%) with size ({widthPercent}%, {heightPercent}%)");

                var activePage = visioApplication?.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[AddShapeToDocument] [Error] No active page found in Visio application.");
                    return;
                }

                // Retrieve page dimensions from Visio
                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                var master = GetShape(categoryName, shapeName);
                if (master == null)
                {
                    Debug.WriteLine($"[AddShapeToDocument] [Error] Shape '{shapeName}' not found in category '{categoryName}'.");
                    return;
                }

                // Calculate scaled coordinates and size
                double visioX = (xPercent / 100.0) * pageWidth;
                double visioY = ((100 - yPercent) / 100.0) * pageHeight;  // Invert Y-axis

                double shapeWidth = (widthPercent / 100.0) * pageWidth;
                double shapeHeight = (heightPercent / 100.0) * pageHeight;

                Debug.WriteLine($"[AddShapeToDocument] Scaled Position - X: {visioX}, Y: {visioY}, Width: {shapeWidth}, Height: {shapeHeight} based on page size: Width={pageWidth}, Height={pageHeight}");

                // Drop shape at calculated coordinates
                var shape = activePage.Drop(master, visioX, visioY);
                shape.Cells["PinX"].ResultIU = visioX;
                shape.Cells["PinY"].ResultIU = visioY;

                // Set shape dimensions explicitly
                shape.Cells["Width"].ResultIU = shapeWidth;
                shape.Cells["Height"].ResultIU = shapeHeight;

                Debug.WriteLine($"[AddShapeToDocument] Shape placed at (PinX={shape.Cells["PinX"].ResultIU}, PinY={shape.Cells["PinY"].ResultIU}) with final size Width={shape.Cells["Width"].ResultIU}, Height={shape.Cells["Height"].ResultIU}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AddShapeToDocument] [Error] Error adding shape '{shapeName}' from category '{categoryName}': {ex.Message}");
                Debug.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
        }

        // New and enhanced functions for greater Visio control:

        public void ConnectShapes(string shape1Name, string shape2Name, string connectorType)
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
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error connecting shapes: {ex.Message}");
            }
        }

        public void AddTextToShape(string shapeName, string text)
        {
            try
            {
                var shape = visioApplication.ActivePage.Shapes.ItemU[shapeName];
                shape.Text = text;
                Debug.WriteLine($"Added text '{text}' to shape: {shapeName}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error adding text to shape: {ex.Message}");
            }
        }

        public void SetShapeStyle(string shapeName, string lineStyle, string fillPattern)
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
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error setting shape style: {ex.Message}");
            }
        }

        public void GroupShapes(string[] shapeNames)
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
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error grouping shapes: {ex.Message}");
            }
        }

        public void UngroupShapes(string shapeName)
        {
            try
            {
                var shape = visioApplication.ActivePage.Shapes.ItemU[shapeName];
                shape.Ungroup();
                Debug.WriteLine($"Ungrouped shape: {shapeName}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error ungrouping shape: {ex.Message}");
            }
        }

        public void AlignShapes(string[] shapeNames, string alignmentType)
        {
            try
            {
                var activePage = visioApplication.ActivePage;
                var selection = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty);
                foreach (var shapeName in shapeNames)
                {
                    selection.Select(activePage.Shapes.ItemU[shapeName], (short)Visio.VisSelectArgs.visSelect);
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
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error aligning shapes: {ex.Message}");
            }
        }

        public void DistributeShapes(string[] shapeNames, string distributionType)
        {
            try
            {
                var activePage = visioApplication.ActivePage;
                var selection = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty);
                foreach (var shapeName in shapeNames)
                {
                    selection.Select(activePage.Shapes.ItemU[shapeName], (short)Visio.VisSelectArgs.visSelect);
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
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error distributing shapes: {ex.Message}");
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

        public List<ShapeInfo> ListAllShapes()
        {
            var shapes = new List<ShapeInfo>();

            try
            {
                var activePage = visioApplication.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[ListAllShapes] [Error] No active page found in Visio.");
                    return shapes;
                }

                foreach (Visio.Shape shape in activePage.Shapes)
                {
                    var shapeInfo = new ShapeInfo
                    {
                        Name = shape.Name,
                        Type = shape.Master?.Name ?? "No Master",
                        Position = new Position
                        {
                            X = shape.CellsU["PinX"].ResultIU,
                            Y = shape.CellsU["PinY"].ResultIU
                        },
                        Color = shape.CellsU["FillForegnd"].FormulaU
                    };

                    shapes.Add(shapeInfo);
                }

                Debug.WriteLine($"[ListAllShapes] Retrieved {shapes.Count} shapes.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ListAllShapes] [Error] Error listing shapes: {ex.Message}");
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
            shapes = new Dictionary<string, Visio.Master>(StringComparer.OrdinalIgnoreCase);
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
            shapes.TryGetValue(name, out Visio.Master master);
            return master;
        }
    }

    public class ShapeInfo
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public Position Position { get; set; }
        public string Color { get; set; }
    }

    public class Position
    {
        public double X { get; set; }
        public double Y { get; set; }
    }
}