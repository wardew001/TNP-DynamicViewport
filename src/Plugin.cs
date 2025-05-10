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
        private Vector2d centerPointOffset = new();
        private double heightOffset = 0;
        private double widthOffset = 0;
        private Point2d lastCenterPoint = new();
        private double lastHeight = 0;
        private double lastWidth = 0;

        public void Initialize()
        {
            Application.DocumentManager.MdiActiveDocument.ViewChanged += OnViewChanged;
            Application.DocumentManager.MdiActiveDocument.LayoutSwitched += OnLayoutSwitched;
        }

        public void Terminate()
        {
            Application.DocumentManager.MdiActiveDocument.ViewChanged -= OnViewChanged;
            Application.DocumentManager.MdiActiveDocument.LayoutSwitched -= OnLayoutSwitched;
        }

        public void OnViewChanged(object sender, EventArgs e)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            using (ViewTableRecord view = ed.GetCurrentView())
            {
                string currentLayout = LayoutManager.Current.CurrentLayout;
                if (currentLayout == "Model")
                {
                    isFirstLoad = false;
                    centerPointOffset += view.CenterPoint - lastCenterPoint;
                    heightOffset += view.Height - lastHeight;
                    widthOffset += view.Width - lastWidth;
                    lastCenterPoint = view.CenterPoint;
                    lastHeight = view.Height;
                    lastWidth = view.Width;
                }
            }
        }

        public void OnLayoutSwitched(object sender, EventArgs e)
        {
            if (isFirstLoad) return;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            using (ViewTableRecord view = ed.GetCurrentView())
            {
                string currentLayout = LayoutManager.Current.CurrentLayout;
                if (currentLayout == "Model")
                {
                    centerPointOffset = new Vector2d(0, 0);
                    heightOffset = 0;
                    widthOffset = 0;
                    lastCenterPoint = view.CenterPoint;
                    lastHeight = view.Height;
                    lastWidth = view.Width;
                }

                if (currentLayout.Contains("Calculations"))
                {
                    ed.WriteMessage("\nSync Action Started.");
                    ed.WriteMessage($"\n[Current] CenterPoint: {view.CenterPoint}, Height: {view.Height}, Width: {view.Width}");
                    using (DocumentLock docLock = doc.LockDocument())
                    {
                        Editor layoutEditor = Application.DocumentManager.MdiActiveDocument.Editor;

                        using (ViewTableRecord layoutView = layoutEditor.GetCurrentView())
                        {
                            layoutView.CenterPoint = view.CenterPoint + centerPointOffset / 100;
                            layoutView.Height = view.Height + heightOffset / 100;
                            layoutView.Width = view.Width + widthOffset / 100;
                            ed.WriteMessage($"\n[New viewport] CenterPoint: {layoutView.CenterPoint}, Height: {layoutView.Height}, Width: {layoutView.Width}");
                            layoutEditor.SetCurrentView(layoutView);
                        }
                    }
                    ed.WriteMessage("\nSync Action Ended.");
                }
            }
        }
    }
}
