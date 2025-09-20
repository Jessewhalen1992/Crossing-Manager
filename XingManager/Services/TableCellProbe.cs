using System;
using System.Collections;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
namespace XingManager.Services
{
    /// <summary>
    /// Reads block attributes from table cells, matching what the Properties palette shows.
    /// Supports both GetBlockAttributeValue overloads and multi-content cells.
    /// </summary>
    public static class TableCellProbe
    {
        public static string TryGetCellBlockAttr(Table t, int row, int col, string tag)
        {
            if (t == null || row < 0 || col < 0 || string.IsNullOrWhiteSpace(tag)) return string.Empty;

            // 1) (row, col, tag, …)
            var v = TryCallGetBlockAttr(t, row, col, tag);
            if (!string.IsNullOrWhiteSpace(v)) return v.Trim();

            // 2) enumerate contents and try (row, col, contentIndex, tag, …)
            var cell = SafeGetCell(t, row, col);
            var contents = GetContents(cell);
            int idx = 0;
            foreach (var _ in contents)
            {
                v = TryCallGetBlockAttrIndexed(t, row, col, idx, tag);
                if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
                idx++;
            }

            // 3) discover tags from the cell's block definition and attempt again
            foreach (var discovered in EnumerateCellBlockTags(t, row, col))
            {
                v = TryCallGetBlockAttr(t, row, col, discovered);
                if (!string.IsNullOrWhiteSpace(v)) return v.Trim();

                idx = 0;
                foreach (var __ in contents)
                {
                    v = TryCallGetBlockAttrIndexed(t, row, col, idx, discovered);
                    if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
                    idx++;
                }
            }

            return string.Empty;
        }

        // --- Diagnostics: read exactly what palette shows for CROSSING in a cell ---
        [CommandMethod("XING_PROBE_CELL")]
        public static void ProbeCell()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed  = doc.Editor;

            var peo = new PromptEntityOptions("\nSelect a TABLE: ");
            peo.AddAllowedClass(typeof(Table), true);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var tbl = (Table)tr.GetObject(per.ObjectId, OpenMode.ForRead);
                var r = ed.GetInteger("\nRow index (0-based): "); if (r.Status != PromptStatus.OK) return;
                var c = ed.GetInteger("\nColumn index (0-based): "); if (c.Status != PromptStatus.OK) return;

                var val = TryGetCellBlockAttr(tbl, r.Value, c.Value, "CROSSING");
                ed.WriteMessage($"\n[CROSSING] => '{val}'");

                tr.Commit();
            }
        }

        // ---------- internals ----------
        private static Cell SafeGetCell(Table t, int row, int col)
        {
            try { return t.Cells[row, col]; } catch { return null; }
        }

        private static IEnumerable GetContents(Cell cell)
        {
            if (cell == null) yield break;
            var p = cell.GetType().GetProperty("Contents", BindingFlags.Public | BindingFlags.Instance);
            var contents = p?.GetValue(cell, null) as IEnumerable;
            if (contents == null) yield break;
            foreach (var c in contents) yield return c;
        }

        private static string TryCallGetBlockAttr(Table t, int row, int col, string tag)
        {
            const string name = "GetBlockAttributeValue";
            foreach (var mi in t.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance))
            {
                if (!string.Equals(mi.Name, name, StringComparison.Ordinal)) continue;
                var p = mi.GetParameters();
                if (p.Length < 3) continue;

                // expect (row, col, tag, …)
                if (typeof(string).IsAssignableFrom(p[2].ParameterType))
                {
                    var args = new object[p.Length];
                    if (!TryConvert(row, p[0], out args[0]) ||
                        !TryConvert(col, p[1], out args[1]) ||
                        !TryConvert(tag, p[2], out args[2])) continue;

                    for (int i = 3; i < p.Length; i++)
                        args[i] = p[i].IsOptional ? Type.Missing : null;

                    try { return Convert.ToString(mi.Invoke(t, args)); } catch { }
                }
            }
            return string.Empty;
        }

        private static string TryCallGetBlockAttrIndexed(Table t, int row, int col, int contentIndex, string tag)
        {
            const string name = "GetBlockAttributeValue";
            foreach (var mi in t.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance))
            {
                if (!string.Equals(mi.Name, name, StringComparison.Ordinal)) continue;
                var p = mi.GetParameters();
                if (p.Length < 4) continue;

                // expect (row, col, contentIndex, tag, …)
                if (typeof(string).IsAssignableFrom(p[3].ParameterType))
                {
                    var args = new object[p.Length];
                    if (!TryConvert(row, p[0], out args[0]) ||
                        !TryConvert(col, p[1], out args[1]) ||
                        !TryConvert(contentIndex, p[2], out args[2]) ||
                        !TryConvert(tag, p[3], out args[3])) continue;

                    for (int i = 4; i < p.Length; i++)
                        args[i] = p[i].IsOptional ? Type.Missing : null;

                    try { return Convert.ToString(mi.Invoke(t, args)); } catch { }
                }
            }
            return string.Empty;
        }

        private static bool TryConvert(object value, ParameterInfo p, out object converted)
        {
            var target = Nullable.GetUnderlyingType(p.ParameterType) ?? p.ParameterType;
            try
            {
                if (value == null) { converted = target.IsValueType ? Activator.CreateInstance(target) : null; return true; }
                if (target.IsInstanceOfType(value)) { converted = value; return true; }
                converted = Convert.ChangeType(value, target, System.Globalization.CultureInfo.InvariantCulture);
                return true;
            }
            catch { converted = null; return false; }
        }

        private static IEnumerable<string> EnumerateCellBlockTags(Table t, int row, int col)
        {
            var cell = SafeGetCell(t, row, col);
            var contents = GetContents(cell);
            var tr = t.Database?.TransactionManager?.TopTransaction as Transaction;
            if (tr == null) yield break;

            foreach (var content in contents)
            {
                var ctProp = content.GetType().GetProperty("ContentTypes", BindingFlags.Public | BindingFlags.Instance);
                var types = ctProp?.GetValue(content, null)?.ToString() ?? string.Empty;
                if (types.IndexOf("Block", StringComparison.OrdinalIgnoreCase) < 0) continue;

                var btrProp = content.GetType().GetProperty("BlockTableRecordId", BindingFlags.Public | BindingFlags.Instance);
                if (!(btrProp?.GetValue(content, null) is ObjectId btrId) || btrId.IsNull || !btrId.IsValid) continue;

                BlockTableRecord btr = null;
                try { btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord; } catch { }
                if (btr == null) continue;

                foreach (ObjectId eid in btr)
                {
                    AttributeDefinition ad = null;
                    try { ad = tr.GetObject(eid, OpenMode.ForRead) as AttributeDefinition; } catch { }
                    if (ad != null && !string.IsNullOrWhiteSpace(ad.Tag)) yield return ad.Tag.Trim();
                }
            }
        }
    }
}
