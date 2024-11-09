using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace VisioPlugin
{
    public class LibraryManager
    {
        private readonly Visio.Application visioApplication;
        private readonly Dictionary<string, ShapeCategory> categories;
        private readonly HttpClient httpClient;
        private readonly string apiEndpoint;

        public LibraryManager(Visio.Application visioApp)
        {
            visioApplication = visioApp ?? throw new ArgumentNullException(nameof(visioApp));
            categories = new Dictionary<string, ShapeCategory>();
            httpClient = new HttpClient();
            apiEndpoint = "http://localhost:5678"; // Set your API endpoint here
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
                Debug.WriteLine($"[Debug] Adding shape: {shapeName} from category: {categoryName} at ({xPercent}%, {yPercent}%) with size ({widthPercent}%, {heightPercent}%)");

                var activePage = visioApplication?.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[Error] No active page found in Visio application.");
                    return;
                }

                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                var master = GetShape(categoryName, shapeName);
                if (master == null)
                {
                    Debug.WriteLine($"[Error] Shape '{shapeName}' not found in category '{categoryName}'.");
                    return;
                }

                double visioX = (xPercent / 100.0) * pageWidth;
                double visioY = (1 - (yPercent / 100.0)) * pageHeight;

                visioX = Math.Max(0, Math.Min(visioX, pageWidth));
                visioY = Math.Max(0, Math.Min(visioY, pageHeight));

                var shape = activePage.Drop(master, visioX, visioY);
                Debug.WriteLine($"[Debug] Shape added: {shape.Name} at ({shape.Cells["PinX"].ResultIU}, {shape.Cells["PinY"].ResultIU})");

                if (widthPercent > 0 && heightPercent > 0)
                {
                    double shapeWidth = (widthPercent / 100.0) * pageWidth;
                    double shapeHeight = (heightPercent / 100.0) * pageHeight;

                    shape.Cells["Width"].ResultIU = Math.Abs(shapeWidth);
                    shape.Cells["Height"].ResultIU = Math.Abs(shapeHeight);

                    Debug.WriteLine($"[Debug] Shape resized: Width = {shape.Cells["Width"].ResultIU}, Height = {shape.Cells["Height"].ResultIU}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Error] Error adding shape '{shapeName}' from category '{categoryName}': {ex.Message}");
                Debug.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
        }

        public void AddShapesToDocument(List<ShapeInfo> shapes)
        {
            try
            {
                var activePage = visioApplication?.ActivePage;
                if (activePage == null)
                {
                    Debug.WriteLine("[Error] No active page found in Visio application.");
                    return;
                }

                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].ResultIU;

                foreach (var shapeInfo in shapes)
                {
                    var master = GetShape(shapeInfo.Category, shapeInfo.Name);
                    if (master == null)
                    {
                        Debug.WriteLine($"[Error] Shape '{shapeInfo.Name}' not found in category '{shapeInfo.Category}'.");
                        continue;
                    }

                    double visioX = (shapeInfo.Position.X / 100.0) * pageWidth;
                    double visioY = (1 - (shapeInfo.Position.Y / 100.0)) * pageHeight;

                    visioX = Math.Max(0, Math.Min(visioX, pageWidth));
                    visioY = Math.Max(0, Math.Min(visioY, pageHeight));

                    var shape = activePage.Drop(master, visioX, visioY);
                    Debug.WriteLine($"[Debug] Shape added: {shape.Name} at ({shape.Cells["PinX"].ResultIU}, {shape.Cells["PinY"].ResultIU})");

                    if (shapeInfo.Size.Width > 0 && shapeInfo.Size.Height > 0)
                    {
                        double shapeWidth = (shapeInfo.Size.Width / 100.0) * pageWidth;
                        double shapeHeight = (shapeInfo.Size.Height / 100.0) * pageHeight;

                        shape.Cells["Width"].ResultIU = Math.Abs(shapeWidth);
                        shape.Cells["Height"].ResultIU = Math.Abs(shapeHeight);

                        Debug.WriteLine($"[Debug] Shape resized: Width = {shape.Cells["Width"].ResultIU}, Height = {shape.Cells["Height"].ResultIU}");
                    }

                    if (!string.IsNullOrEmpty(shapeInfo.Color))
                    {
                        SetShapeColor(shape, shapeInfo.Color);
                        Debug.WriteLine($"[Debug] Shape color set to: {shapeInfo.Color}");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Error] Error adding shapes: {ex.Message}");
                Debug.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
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

        public async Task SendLibraryInformationToN8n()
        {
            try
            {
                var libraryInfo = ListAllShapes();
                var jsonContent = new StringContent(JsonConvert.SerializeObject(libraryInfo), Encoding.UTF8, "application/json");

                var response = await httpClient.PostAsync($"{apiEndpoint}/send-library-info", jsonContent);
                response.EnsureSuccessStatusCode();

                Debug.WriteLine("Library information sent to n8n successfully.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error sending library information to n8n: {ex.Message}");
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
        public string Category { get; set; }
        public Size Size { get; set; }
    }

    public class Position
    {
        public double X { get; set; }
        public double Y { get; set; }
    }

    public class Size
    {
        public double Width { get; set; }
        public double Height { get; set; }
    }
}
