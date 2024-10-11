using System;
using System.Collections.Generic;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioPlugin
{
    public class VisioAutomation
    {
        private readonly Visio.Application _visioApplication;
        private Visio.Document _activeDocument;
        private Visio.Page _activePage;
        private Dictionary<string, Visio.Document> _loadedStencils;

        public VisioAutomation(Visio.Application visioApplication)
        {
            _visioApplication = visioApplication;
            _loadedStencils = new Dictionary<string, Visio.Document>();
            UpdateActiveDocumentAndPage();
        }

        private void UpdateActiveDocumentAndPage()
        {
            _activeDocument = _visioApplication.ActiveDocument;
            _activePage = _activeDocument?.Pages.Cast<Visio.Page>().FirstOrDefault(p => p.ID == _activeDocument.Application.ActivePage.ID);
        }

        public void LoadStencil(string stencilPath, string stencilName)
        {
            try
            {
                if (!_loadedStencils.ContainsKey(stencilName))
                {
                    var stencil = _visioApplication.Documents.OpenEx(stencilPath, (short)Visio.VisOpenSaveArgs.visOpenDocked);
                    _loadedStencils[stencilName] = stencil;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to load stencil: {ex.Message}", ex);
            }
        }

        public void AddShapeFromStencil(string stencilName, string masterName, double x, double y)
        {
            UpdateActiveDocumentAndPage();
            if (_activePage == null)
            {
                throw new InvalidOperationException("No active page found.");
            }

            if (!_loadedStencils.TryGetValue(stencilName, out var stencil))
            {
                throw new ArgumentException($"Stencil '{stencilName}' not loaded.");
            }

            Visio.Master master = stencil.Masters[masterName];
            if (master == null)
            {
                throw new ArgumentException($"Master '{masterName}' not found in the stencil '{stencilName}'.");
            }

            _activePage.Drop(master, x, y);
        }

        public List<string> GetLoadedStencils()
        {
            return new List<string>(_loadedStencils.Keys);
        }

        public List<string> GetMastersInStencil(string stencilName)
        {
            if (!_loadedStencils.TryGetValue(stencilName, out var stencil))
            {
                throw new ArgumentException($"Stencil '{stencilName}' not loaded.");
            }

            var masters = new List<string>();
            foreach (Visio.Master master in stencil.Masters)
            {
                masters.Add(master.Name);
            }
            return masters;
        }

        public void AddShape(string masterName, double x, double y)
        {
            UpdateActiveDocumentAndPage();
            if (_activePage == null)
            {
                throw new InvalidOperationException("No active page found.");
            }

            Visio.Master master = _activeDocument.Masters[masterName];
            if (master == null)
            {
                throw new ArgumentException($"Master '{masterName}' not found in the active document.");
            }

            _activePage.Drop(master, x, y);
        }

        public void ConnectShapes(Visio.Shape shape1, Visio.Shape shape2)
        {
            UpdateActiveDocumentAndPage();
            if (_activePage == null)
            {
                throw new InvalidOperationException("No active page found.");
            }

            Visio.Shape connector = _activePage.Drop(_activeDocument.Masters["Dynamic connector"], 0, 0);
            connector.CellsU["BeginX"].GlueTo(shape1.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowXFormOut, (short)Visio.VisCellIndices.visXFormPinX]);
            connector.CellsU["EndX"].GlueTo(shape2.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowXFormOut, (short)Visio.VisCellIndices.visXFormPinX]);
        }

        public void HighlightShape(Visio.Shape shape, System.Drawing.Color color)
        {
            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineColor].FormulaU = $"RGB({color.R},{color.G},{color.B})";
            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineWeight].FormulaU = "2 pt";
        }

        public void ResetShapeHighlight(Visio.Shape shape)
        {
            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineColor].FormulaU = "Guard(0)";
            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineWeight].FormulaU = "Guard(0)";
        }
    }
}