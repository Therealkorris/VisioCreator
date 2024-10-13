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
                    }
                }
            }
        }

        private void LoadAllStencilDocuments()
        {
            foreach (Visio.Document doc in visioApplication.Documents)
            {
                if (doc.Type == Visio.VisDocumentTypes.visTypeStencil)
                {
                    string category = doc.Name;
                    if (!categories.ContainsKey(category))
                    {
                        categories[category] = new ShapeCategory(category);
                    }

                    foreach (Visio.Master master in doc.Masters)
                    {
                        categories[category].AddShape(master.Name, master);
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
            return categories.Keys;
        }

        public IEnumerable<string> GetShapesInCategory(string categoryName)
        {
            var shapes = new List<string>();
            var stencil = visioApplication.Documents.OpenEx(categoryName, (short)Visio.VisOpenSaveArgs.visOpenDocked);

            foreach (Visio.Master master in stencil.Masters)
            {
                shapes.Add(master.Name);
            }

            return shapes;
        }

        public void AddShapeToDocument(string categoryName, string shapeName, double x, double y, double width, double height)
        {
            Debug.WriteLine($"Adding shape: {shapeName} from category: {categoryName} at ({x}, {y}) with size ({width}, {height})");

            var activePage = visioApplication.ActivePage;

            // Open the stencil to get the master (shape template)
            var stencil = visioApplication.Documents.OpenEx(categoryName, (short)Visio.VisOpenSaveArgs.visOpenDocked);
            var master = stencil.Masters[shapeName];

            // Drop the shape on the active Visio page
            var shape = activePage.Drop(master, x, y);

            Debug.WriteLine($"Shape added: {shape.Name} at ({shape.Cells["PinX"].ResultIU}, {shape.Cells["PinY"].ResultIU})");

            // Adjust size if provided
            shape.Cells["Width"].ResultIU = width;
            shape.Cells["Height"].ResultIU = height;

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
