using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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
                        Type = shape.Master.Name,
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
