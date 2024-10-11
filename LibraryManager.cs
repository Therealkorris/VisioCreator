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
                        Debug.WriteLine($"Added shape: {master.Name} to category: {category}");
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
                    Debug.WriteLine($"Loading stencil: {doc.Name}");
                    string category = doc.Name;
                    if (!categories.ContainsKey(category))
                    {
                        categories[category] = new ShapeCategory(category);
                    }

                    foreach (Visio.Master master in doc.Masters)
                    {
                        categories[category].AddShape(master.Name, master);
                        Debug.WriteLine($"Added shape: {master.Name} to category: {category}");
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

        public IEnumerable<string> GetShapesInCategory(string category)
        {
            if (categories.TryGetValue(category, out ShapeCategory shapeCategory))
            {
                return shapeCategory.GetShapeNames();
            }

            return Enumerable.Empty<string>();
        }

        public void AddShapeToDocument(string category, string shapeName, double x, double y)
        {
            if (categories.TryGetValue(category, out ShapeCategory shapeCategory))
            {
                Visio.Master master = shapeCategory.GetShape(shapeName);
                if (master != null)
                {
                    visioApplication.ActivePage.Drop(master, x, y);
                }
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
            shapes.TryGetValue(name, out Visio.Master master);
            return master;
        }
    }
}
