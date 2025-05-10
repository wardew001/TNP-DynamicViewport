using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;

[assembly: CommandClass(typeof(TNPPlugin.Plugin))]

namespace TNPPlugin
{
    public class Plugin : IExtensionApplication
    {
        private bool isFirstLoad = true;
        private Point2d lastCenterPoint = new();
        private double lastHeight = 0;
        private double lastWidth = 0;

        public void Initialize()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.ViewChanged += OnViewChanged;
                doc.LayoutSwitched += OnLayoutSwitched;
            }
        }

        public void Terminate()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.ViewChanged -= OnViewChanged;
                doc.LayoutSwitched -= OnLayoutSwitched;
            }
        }

        public void OnViewChanged(object sender, EventArgs e)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            using var view = ed.GetCurrentView();
            string layout = LayoutManager.Current.CurrentLayout;

            if (layout != "Model") return;

            isFirstLoad = false;

            lastCenterPoint = view.CenterPoint;
            lastHeight = view.Height;
            lastWidth = view.Width;
        }

        public void OnLayoutSwitched(object sender, EventArgs e)
        {
            if (isFirstLoad) return;

            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            using var view = ed.GetCurrentView();
            string layout = LayoutManager.Current.CurrentLayout;

            if (layout == "Model")
            {
                lastCenterPoint = view.CenterPoint;
                lastHeight = view.Height;
                lastWidth = view.Width;
                return;
            }

            if (!layout.Contains("Calculations")) return;

            ed.WriteMessage("\nSync Action Started.");

            using (var docLock = doc.LockDocument())
            {
                AdjustTargetViewport(doc, ed, view);
            }

            ed.WriteMessage("\nSync Action Ended.");
        }

        private void AdjustTargetViewport(Document doc, Editor ed, ViewTableRecord modelView)
        {
            Database db = doc.Database;

            using var tr = db.TransactionManager.StartTransaction();
            var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            if (!layoutDict.Contains("36 Detention Pond Layout and Calculations"))
            {
                ed.WriteMessage("\nLayout 'Calculations' not found.");
                return;
            }

            var layoutId = layoutDict.GetAt("36 Detention Pond Layout and Calculations");
            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
            var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

            foreach (ObjectId entId in btr)
            {
                if (tr.GetObject(entId, OpenMode.ForWrite) is Viewport vp && vp.Number > 1 && !vp.IsErased)
                {
                    ed.WriteMessage($"\n[Current] CenterPoint: {vp.ViewCenter}, Height: {vp.ViewHeight}, Width: {vp.CustomScale}");

                    vp.ViewCenter = lastCenterPoint;
                    vp.ViewHeight = lastHeight;

                    ed.WriteMessage($"\n[Viewport Updated] ViewCenter: {vp.ViewCenter}, ViewHeight: {vp.ViewHeight}, CustomScale: {vp.CustomScale}");
                    break;
                }
            }

            tr.Commit();
        }
    }
}
