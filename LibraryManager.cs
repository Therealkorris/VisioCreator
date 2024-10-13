using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioPlugin
{
    public class LibraryManager
    {
        private readonly Visio.Application visioApplication;
        private readonly Dictionary<string, ShapeCategory> categories;

        public LibraryManager(Visio.Application visioApp)
        {
            visioApplication = visioApp;
            categories = new Dictionary<string, ShapeCategory>();
        }

        public void LoadLibraries()
        {
            categories.Clear();

            // Scan all stencils that are currently open
            ScanAvailableShapes();

            // Load all accessible stencil documents from Visio
            LoadAllStencilDocuments();
        }

        private void ScanAvailableShapes()
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
                        //Debug.WriteLine($"Added shape: {master.Name} to category: {category}");
                    }
                }
            }
        }

        private void LoadAllStencilDocuments()
        {
            // Get all open stencil documents (this includes the built-in and user-added stencils)
            foreach (Visio.Document doc in visioApplication.Documents)
            {
                if (doc.Type == Visio.VisDocumentTypes.visTypeStencil)
                {
                    //Debug.WriteLine($"Loading stencil: {doc.Name}");
                    string category = doc.Name;
                    if (!categories.ContainsKey(category))
                    {
                        categories[category] = new ShapeCategory(category);
                    }

                    foreach (Visio.Master master in doc.Masters)
                    {
                        categories[category].AddShape(master.Name, master);
                        //Debug.WriteLine($"Added shape: {master.Name} to category: {category}");
                    }
                }
            }
        }

        public IEnumerable<string> GetCategories()
        {
            if (!categories.Any())
            {
                LoadLibraries();
            }
            return categories.Keys.ToList();
        }

        public IEnumerable<string> GetShapesInCategory(string categoryName)
        {
            // Assuming we have a way to load the library
            var shapes = new List<string>();

            // Example of loading shapes from a specific stencil or category
            var stencil = visioApplication.Documents.OpenEx(categoryName, (short)Visio.VisOpenSaveArgs.visOpenDocked);

            foreach (Visio.Master master in stencil.Masters)
            {
                shapes.Add(master.Name);
            }

            return shapes;
        }

        public void AddShapeToDocument(string categoryName, string shapeName, double x, double y)
        {
            Debug.WriteLine($"Adding shape: {shapeName} from category: {categoryName} at ({x}, {y})");
            
            // Open the stencil that contains the shapes
            var stencil = visioApplication.Documents.OpenEx(categoryName, (short)Visio.VisOpenSaveArgs.visOpenDocked);

            // Find the master shape by name
            var master = stencil.Masters[shapeName];

            // Add the shape to the active page at the specified position
            var activePage = visioApplication.ActivePage;
            var shape = activePage.Drop(master, x / 100, y / 100);  // Divide by 100 to scale down

            Debug.WriteLine($"Shape added: {shape.Name} at ({shape.Cells["PinX"].ResultIU}, {shape.Cells["PinY"].ResultIU})");
            
            // Set a fixed size for the shape
            shape.Cells["Width"].ResultIU = 1;  // 1 inch width
            shape.Cells["Height"].ResultIU = 1;  // 1 inch height

            Debug.WriteLine($"Shape resized: Width = {shape.Cells["Width"].ResultIU}, Height = {shape.Cells["Height"].ResultIU}");
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
            shapes.TryGetValue(name, out Visio.Master master);
            return master;
        }
    }
}
