using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

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
    }
}
