using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using XingManager.Services;

namespace XingManager
{
    public class DebugCommands
    {
        // Add this to your main commands class file
        [CommandMethod("XING_DUMP_CELL")]
        public void XingDumpCell()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            // Prompt for the table
            var peo = new PromptEntityOptions("\nSelect a TABLE:");
            peo.SetRejectMessage("\nMust be a Table object.");
            peo.AddAllowedClass(typeof(Table), true);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            // Prompt for row index
            var pioRow = new PromptIntegerOptions("\nRow index (0-based):");
            pioRow.AllowNegative = false;
            var pirRow = ed.GetInteger(pioRow);
            if (pirRow.Status != PromptStatus.OK) return;
            int row = pirRow.Value;

            // Prompt for column index
            var pioCol = new PromptIntegerOptions("\nColumn index (0-based):");
            pioCol.AllowNegative = false;
            var pirCol = ed.GetInteger(pioCol);
            if (pirCol.Status != PromptStatus.OK) return;
            int col = pirCol.Value;

            ed.WriteMessage($"\n[Dump] Cell ({row},{col})\n");

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var table = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table;
                if (table == null) return;

                if (row >= table.Rows.Count || col >= table.Columns.Count)
                {
                    ed.WriteMessage("\nError: Row or column index is out of bounds.");
                    return;
                }

                var cell = table.Cells[row, col];

                ed.WriteMessage("\n--- Cell Properties ---\n");
                ed.WriteMessage($"Cell.DataType: {cell.DataType}\n");
                ed.WriteMessage($"Cell.Value Type: {(cell.Value?.GetType().FullName ?? "null")}\n");
                ed.WriteMessage($"Cell.Value: {(cell.Value?.ToString() ?? "null")}\n");
                ed.WriteMessage($"Cell.TextString: {cell.TextString}\n");

                if (cell.Value is ObjectId blockId && !blockId.IsNull)
                {
                    ed.WriteMessage("\n--- Block Attributes ---\n");
                    var blockRef = tr.GetObject(blockId, OpenMode.ForRead) as BlockReference;
                    if (blockRef != null)
                    {
                        ed.WriteMessage($"Block Name: {blockRef.Name}\n");
                        if (blockRef.AttributeCollection.Count == 0)
                        {
                            ed.WriteMessage("-> No attributes found on this block.\n");
                        }
                        foreach (ObjectId attId in blockRef.AttributeCollection)
                        {
                            var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                            if (attRef != null)
                            {
                                ed.WriteMessage($"-> TAG: '{attRef.Tag}' = VALUE: '{attRef.TextString}'\n");
                            }
                        }
                    }
                    else
                    {
                        ed.WriteMessage("-> Value is an ObjectId, but it's not a BlockReference.\n");
                    }
                }
                else
                {
                    ed.WriteMessage("\n-> Cell.Value is NOT a block ObjectId.\n");
                }
                tr.Commit();
            }
        }

        [CommandMethod("XING_FIXBORDERS")]
        public void FixAllTableBorders()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (doc.LockDocument())
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    foreach (ObjectId entId in btr)
                    {
                        var t = tr.GetObject(entId, OpenMode.ForRead) as Table;
                        if (t == null) continue;
                        t.UpgradeOpen();

                        int rows = t.Rows.Count;
                        int cols = t.Columns.Count;

                        int dataStart = 0;
                        try { dataStart = TableSync.FindLatLongDataStartRow(t); if (dataStart < 0) dataStart = 0; } catch { dataStart = 0; }

                        bool IsHeading(int r)
                        {
                            try
                            {
                                var s = (t.Cells[r, 0]?.TextString ?? string.Empty).Trim().ToUpperInvariant();
                                if (s.Contains("CROSSING INFORMATION")) return true;
                            }
                            catch { }
                            return (dataStart > 0 && r == dataStart - 1);
                        }

                        for (int r = 0; r < rows; r++)
                            for (int c = 0; c < cols; c++)
                            {
                                bool heading = IsHeading(r);
                                TrySetGridVisibility(t, r, c, GridLineType.HorizontalTop, !heading);
                                TrySetGridVisibility(t, r, c, GridLineType.VerticalRight, !heading);
                                TrySetGridVisibility(t, r, c, GridLineType.HorizontalBottom, true);
                                TrySetGridVisibility(t, r, c, GridLineType.VerticalLeft, !heading);
                            }

                        try { t.GenerateLayout(); } catch { }
                        try { t.RecordGraphicsModified(true); } catch { }
                    }
                }
                tr.Commit();
            }

            try { doc.Editor?.Regen(); } catch { }
        }

        private static MethodInfo _debugSetGridVisibilityBool;
        private static MethodInfo _debugSetGridVisibilityEnum;
        private static Type _debugVisibilityType;

        private static void TrySetGridVisibility(Table t, int row, int col, GridLineType line, bool visible)
        {
            if (t == null) return;

            try
            {
                var tableType = t.GetType();
                if (_debugSetGridVisibilityBool == null)
                {
                    _debugSetGridVisibilityBool = tableType.GetMethod(
                        "SetGridVisibility",
                        new[] { typeof(int), typeof(int), typeof(GridLineType), typeof(bool) });
                }

                if (_debugSetGridVisibilityBool != null)
                {
                    _debugSetGridVisibilityBool.Invoke(t, new object[] { row, col, line, visible });
                    return;
                }

                if (_debugVisibilityType == null)
                {
                    _debugVisibilityType = tableType.Assembly.GetType("Autodesk.AutoCAD.DatabaseServices.Visibility");
                }
                if (_debugVisibilityType == null)
                    return;

                if (_debugSetGridVisibilityEnum == null)
                {
                    _debugSetGridVisibilityEnum = tableType.GetMethod(
                        "SetGridVisibility",
                        new[] { typeof(int), typeof(int), typeof(GridLineType), _debugVisibilityType });
                }

                if (_debugSetGridVisibilityEnum == null)
                    return;

                object enumValue = null;
                try
                {
                    var field = _debugVisibilityType.GetField(visible ? "Visible" : "Invisible", BindingFlags.Public | BindingFlags.Static);
                    enumValue = field?.GetValue(null);
                }
                catch
                {
                    enumValue = null;
                }

                if (enumValue == null)
                {
                    try
                    {
                        enumValue = Enum.Parse(_debugVisibilityType, visible ? "Visible" : "Invisible");
                    }
                    catch
                    {
                        return;
                    }
                }

                _debugSetGridVisibilityEnum.Invoke(t, new[] { (object)row, (object)col, (object)line, enumValue });
            }
            catch
            {
                // best effort compatibility for different AutoCAD versions
            }
        }
    }
}
