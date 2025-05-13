using System;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(TNPPlugin.Plugin))]

namespace TNPPlugin
{
    public class Plugin : IExtensionApplication
    {
        private const string TargetLayoutName = "36 Detention Pond Layout and Calculations";
        private const string RegAppName = "TNPPlugin";

        public void Initialize()
        {
            Application.DocumentManager.MdiActiveDocument.LayoutSwitched += OnLayoutSwitched;

            var db = HostApplicationServices.WorkingDatabase;
            using var tr = db.TransactionManager.StartTransaction();
            var rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
            if (!rat.Has(RegAppName))
            {
                rat.UpgradeOpen();
                var rec = new RegAppTableRecord { Name = RegAppName };
                rat.Add(rec);
                tr.AddNewlyCreatedDBObject(rec, true);
            }
            tr.Commit();
        }

        public void Terminate()
        {
            Application.DocumentManager.MdiActiveDocument.LayoutSwitched -= OnLayoutSwitched;
        }

        [CommandMethod("TNP_SELECT_AREA")]
        public void CreateArea()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            if (!LayoutManager.Current.CurrentLayout.Equals("Model", StringComparison.OrdinalIgnoreCase))
            {
                ed.WriteMessage("\nSwitch to Model layout first.");
                return;
            }

            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                foreach (ObjectId id in ms)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Polyline;
                    if (ent != null && ent.XData != null)
                    {
                        var data = ent.XData.AsArray();
                        if (data.Length > 0 &&
                            data[0].TypeCode == (int)DxfCode.ExtendedDataRegAppName &&
                            data[0].Value.ToString() == "TNPPlugin")
                        {
                            ent.UpgradeOpen();
                            ent.Erase();
                        }
                    }
                }
                tr.Commit();
            }

            var p1res = ed.GetPoint("\nFirst corner of area: ");
            if (p1res.Status != PromptStatus.OK) return;
            var p2res = ed.GetCorner(new PromptCornerOptions("\nOpposite corner: ", p1res.Value));
            if (p2res.Status != PromptStatus.OK) return;

            var ucs2wcs = ed.CurrentUserCoordinateSystem;
            var wcsP1 = p1res.Value.TransformBy(ucs2wcs);
            var wcsP2 = p2res.Value.TransformBy(ucs2wcs);

            var minPt = new Point3d(Math.Min(wcsP1.X, wcsP2.X),
                                    Math.Min(wcsP1.Y, wcsP2.Y),
                                    0.0);
            var maxPt = new Point3d(Math.Max(wcsP1.X, wcsP2.X),
                                    Math.Max(wcsP1.Y, wcsP2.Y),
                                    0.0);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var rect = new Polyline();
                rect.AddVertexAt(0, new Point2d(minPt.X, minPt.Y), 0, 0, 0);
                rect.AddVertexAt(1, new Point2d(maxPt.X, minPt.Y), 0, 0, 0);
                rect.AddVertexAt(2, new Point2d(maxPt.X, maxPt.Y), 0, 0, 0);
                rect.AddVertexAt(3, new Point2d(minPt.X, maxPt.Y), 0, 0, 0);
                rect.Closed = true;

                var xdata = new ResultBuffer(
                    new TypedValue((int)DxfCode.ExtendedDataRegAppName, RegAppName),
                    new TypedValue((int)DxfCode.ExtendedDataInteger16, 1)
                );
                rect.XData = xdata;

                btr.AppendEntity(rect);
                tr.AddNewlyCreatedDBObject(rect, true);
                tr.Commit();
            }

            ed.WriteMessage($"\nArea stored and rectangle created. Will fit into layout '{TargetLayoutName}'.");
        }

        private void OnLayoutSwitched(object sender, EventArgs e)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            string layoutName = LayoutManager.Current.CurrentLayout;
            if (!layoutName.Equals(TargetLayoutName, StringComparison.OrdinalIgnoreCase))
                return;

            Extents3d? extents = null;

            using var tr = db.TransactionManager.StartTransaction();
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId entId in btr)
            {
                if (tr.GetObject(entId, OpenMode.ForRead) is Polyline pline
                    && pline.Closed
                    && pline.XData != null)
                {
                    var xdata = pline.XData;
                    if (xdata.AsArray().Any(tv => tv.TypeCode == (int)DxfCode.ExtendedDataRegAppName
                                              && tv.Value.ToString() == RegAppName))
                    {
                        extents = pline.GeometricExtents;
                        break;
                    }
                }
            }

            if (extents == null)
            {
                ed.WriteMessage("\nNo area rectangle found; run TNP_SELECT_AREA first.");
                return;
            }

            var layDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
            if (!layDict.Contains(TargetLayoutName))
            {
                ed.WriteMessage($"\nLayout '{TargetLayoutName}' not found.");
                return;
            }

            var layId = layDict.GetAt(TargetLayoutName);
            var layout = (Layout)tr.GetObject(layId, OpenMode.ForRead);
            var layoutBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

            foreach (ObjectId entId in layoutBtr)
            {
                if (tr.GetObject(entId, OpenMode.ForWrite) is Viewport vp
                    && vp.Number > 1
                    && !vp.IsErased)
                {
                    vp.ZoomToExtents3d(extents.Value);
                    ed.WriteMessage("\nViewport zoomed to fit stored area.");
                    break;
                }
            }

            tr.Commit();
        }
    }

    public static class ViewportExtensions
    {
        public static void ZoomToExtents3d(this Viewport vp, Extents3d extents)
        {
            var minPt = extents.MinPoint;
            var maxPt = extents.MaxPoint;
            var corners = new[]
            {
                new Point3d(minPt.X, minPt.Y, minPt.Z),
                new Point3d(maxPt.X, minPt.Y, minPt.Z),
                new Point3d(minPt.X, maxPt.Y, minPt.Z),
                new Point3d(maxPt.X, maxPt.Y, minPt.Z),
                new Point3d(minPt.X, minPt.Y, maxPt.Z),
                new Point3d(maxPt.X, minPt.Y, maxPt.Z),
                new Point3d(minPt.X, maxPt.Y, maxPt.Z),
                new Point3d(maxPt.X, maxPt.Y, maxPt.Z)
            };

            var wcs2dcs = Matrix3d.WorldToPlane(vp.ViewDirection)
                        * Matrix3d.Displacement(vp.ViewTarget.GetAsVector().Negate())
                        * Matrix3d.Rotation(vp.TwistAngle, vp.ViewDirection, vp.ViewTarget);

            var dcsExt = corners
                .Select(pt => pt.TransformBy(wcs2dcs))
                .Aggregate(new Extents3d(),
                    (agg, p) => { agg.AddPoint(p); return agg; });

            var dcsCenter = new LineSegment3d(dcsExt.MinPoint, dcsExt.MaxPoint).MidPoint;

            double vpRatio = vp.Width / vp.Height;
            double extWidth = dcsExt.MaxPoint.X - dcsExt.MinPoint.X;
            double extHeight = dcsExt.MaxPoint.Y - dcsExt.MinPoint.Y;
            double scale = (extWidth / extHeight) < vpRatio
                             ? vp.Height / extHeight
                             : vp.Width / extWidth;

            if (!vp.IsWriteEnabled) vp.UpgradeOpen();
            vp.ViewCenter = new Point2d(dcsCenter.X, dcsCenter.Y);
            vp.CustomScale = scale;
        }
    }
}
