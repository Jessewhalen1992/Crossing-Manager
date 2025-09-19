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
        [CommandMethod("XING_DUMP_CELL")]
        public void DumpTableCell()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed  = doc.Editor;

            try
            {
                // 1) Select a Table
                var peo = new PromptEntityOptions("\nSelect a TABLE:");
                peo.SetRejectMessage("\nMust be a Table.");
                peo.AddAllowedClass(typeof(Table), exactMatch: true);
                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                // 2) Row/col (0-based)
                var prRow = ed.GetInteger("\nRow index (0-based): ");
                if (prRow.Status != PromptStatus.OK) return;
                var prCol = ed.GetInteger("\nColumn index (0-based): ");
                if (prCol.Status != PromptStatus.OK) return;

                using (doc.LockDocument())
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var tbl = (Table)tr.GetObject(per.ObjectId, OpenMode.ForRead);

                    if (prRow.Value < 0 || prRow.Value >= tbl.Rows.Count ||
                        prCol.Value < 0 || prCol.Value >= tbl.Columns.Count)
                    {
                        ed.WriteMessage("\n[Dump] Row/Col out of range.");
                        return;
                    }

                    var cell = tbl.Cells[prRow.Value, prCol.Value];
                    ed.WriteMessage($"\n[Dump] Cell ({prRow.Value},{prCol.Value})");

                    // A) Direct text
                    try { ed.WriteMessage($"\n  TextString: '{(cell.TextString ?? string.Empty).Trim()}'"); }
                    catch { ed.WriteMessage("\n  TextString: <error>"); }

                    // B) Table-level block attribute (try common tags)
                    var guessTags = new[] { "CROSSING","XING","X_NO","XNUM","XNUMBER","NUMBER","INDEX","NO","LABEL" };
                    var getM = typeof(Table).GetMethods().FirstOrDefault(m => m.Name == "GetBlockAttributeValue" && m.GetParameters().Length >= 3);
                    if (getM != null)
                    {
                        foreach (var tag in guessTags)
                        {
                            try
                            {
                                var pars = getM.GetParameters();
                                var args = new object[pars.Length];
                                args[0] = Convert.ChangeType(prRow.Value, pars[0].ParameterType, CultureInfo.InvariantCulture);
                                args[1] = Convert.ChangeType(prCol.Value, pars[1].ParameterType, CultureInfo.InvariantCulture);
                                args[2] = tag;
                                for (int i = 3; i < pars.Length; i++) args[i] = Type.Missing;

                                var val = getM.Invoke(tbl, args) as string;
                                ed.WriteMessage($"\n  GetBlockAttributeValue('{tag}') => '{(val ?? string.Empty).Trim()}'");
                            }
                            catch { ed.WriteMessage($"\n  GetBlockAttributeValue('{tag}') => <error>"); }
                        }
                    }
                    else
                    {
                        ed.WriteMessage("\n  GetBlockAttributeValue: not available on this Table version.");
                    }

                    // C) Cell-level property BlockAttributeValue
                    try
                    {
                        var p = cell.GetType().GetProperty("BlockAttributeValue", BindingFlags.Public | BindingFlags.Instance);
                        if (p != null)
                        {
                            var v = p.GetValue(cell, null) as string;
                            ed.WriteMessage($"\n  Cell.BlockAttributeValue => '{(v ?? string.Empty).Trim()}'");
                        }
                        else ed.WriteMessage("\n  Cell.BlockAttributeValue: <not present>");
                    }
                    catch { ed.WriteMessage("\n  Cell.BlockAttributeValue => <error>"); }

                    // D) Dump all public props on each content item (if any)
                    try
                    {
                        var contentsProp = cell.GetType().GetProperty("Contents", BindingFlags.Public | BindingFlags.Instance);
                        if (contentsProp != null)
                        {
                            var contents = contentsProp.GetValue(cell, null) as System.Collections.IEnumerable;
                            if (contents != null)
                            {
                                int i = 0;
                                foreach (var item in contents)
                                {
                                    ed.WriteMessage($"\n  Content[{i}] Type = {item?.GetType().FullName}");
                                    foreach (var prop in item.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
                                    {
                                        object pv = null;
                                        try { pv = prop.GetValue(item, null); } catch {}
                                        var s = pv is string ? (string)pv : pv?.ToString() ?? "";
                                        if (!string.IsNullOrEmpty(s))
                                            ed.WriteMessage($"\n    {prop.Name} = '{s.Trim()}'");
                                    }
                                    i++;
                                }
                                if (i == 0) ed.WriteMessage("\n  Contents: (none)");
                            }
                            else ed.WriteMessage("\n  Contents: <not enumerable>");
                        }
                        else ed.WriteMessage("\n  Contents: <property not present>");
                    }
                    catch { ed.WriteMessage("\n  Contents dump failed."); }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[Dump] Error: {ex.Message}");
            }
        }
    }
}
