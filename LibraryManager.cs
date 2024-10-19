using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioPlugin
{
    public class LibraryManager
    {
        private readonly Visio.Application visioApplication;
        private readonly Dictionary<string, ShapeCategory> categories;
        private string stencilPath;

        public LibraryManager(Visio.Application visioApp)
        {
            visioApplication = visioApp;
            categories = new Dictionary<string, ShapeCategory>();
            LoadLibraries();
        }

        public void LoadLibraries()
        {
            categories.Clear();

            // Load the specific stencil: BASIC_M.vssx
            LoadSpecificStencil();

            // Build the shapes catalog
            BuildShapesCatalog();
        }

        private void LoadSpecificStencil()
        {
            try
            {
                // Replace this path with the actual location of your BASIC_M.vssx file
                stencilPath = @"C:\Users\%username%\Documents\My Shapes\BASIC_M.vssx";

                if (File.Exists(stencilPath))
                {
                    visioApplication.Documents.OpenEx(stencilPath,
                        (short)Visio.VisOpenSaveArgs.visOpenHidden |
                        (short)Visio.VisOpenSaveArgs.visOpenRO);

                    Debug.WriteLine($"Loaded stencil: {stencilPath}");
                }
                else
                {
                    Debug.WriteLine($"Stencil not found at path: {stencilPath}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading specific stencil: {ex.Message}");
            }
        }

        private void BuildShapesCatalog()
        {
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

        public string FindCategoryForShape(string shapeName)
        {
            foreach (var category in categories.Values)
            {
                if (category.ContainsShape(shapeName))
                {
                    return category.Name;
                }
            }
            return null;
        }

        public void AddShapeToDocument(string categoryName, string shapeName, double x, double y, double width, double height)
        {
            try
            {
                Debug.WriteLine($"Adding shape: {shapeName} from category: {categoryName} at ({x}, {y}) with size ({width}, {height})");

                var activePage = visioApplication.ActivePage;

                // Open the stencil to get the master (shape template)
                var stencil = visioApplication.Documents.OpenEx(stencilPath, (short)Visio.VisOpenSaveArgs.visOpenHidden | (short)Visio.VisOpenSaveArgs.visOpenRO);
                var master = stencil.Masters.get_ItemU(shapeName);

                if (master == null)
                {
                    Debug.WriteLine($"Shape '{shapeName}' not found in stencil '{stencil.Name}'.");
                    return;
                }

                // Drop the shape on the active Visio page
                var shape = activePage.Drop(master, x, y);

                Debug.WriteLine($"Shape added: {shape.Name} at ({shape.Cells["PinX"].ResultIU}, {shape.Cells["PinY"].ResultIU})");

                // Adjust size if provided
                if (width > 0 && height > 0)
                {
                    shape.Cells["Width"].ResultIU = width;
                    shape.Cells["Height"].ResultIU = height;
                }

                Debug.WriteLine($"Shape resized: Width = {shape.Cells["Width"].ResultIU}, Height = {shape.Cells["Height"].ResultIU}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error adding shape '{shapeName}' from stencil '{stencilPath}': {ex.Message}");
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
                shape.CellsU["FillPattern"].FormulaU = "1"; // Ensure the shape has a fill

                Debug.WriteLine($"Set color for shape '{shape.Name}' to {colorHex}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error setting color for shape '{shape.Name}': {ex.Message}");
            }
        }

        // ... (Rest of your code)
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

        public bool ContainsShape(string name)
        {
            return shapes.ContainsKey(name);
        }
    }
}
