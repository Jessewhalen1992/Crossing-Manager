using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using XingManager.Models;

namespace XingManager.Services
{
    public class TableMatcher
    {
        private static readonly ISet<string> CrossingAttributeTags = new HashSet<string>(new[]
        {
            "CROSSING",
            "XING",
            "X_NO",
            "XNUM",
            "XNUMBER",
            "NUMBER",
            "INDEX",
            "NO",
            "LABEL"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> OwnerAttributeTags = new HashSet<string>(new[]
        {
            "OWNER",
            "OWN",
            "COMPANY"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> DescriptionAttributeTags = new HashSet<string>(new[]
        {
            "DESCRIPTION",
            "DESC"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> LocationAttributeTags = new HashSet<string>(new[]
        {
            "LOCATION",
            "LOC"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> DwgRefAttributeTags = new HashSet<string>(new[]
        {
            "DWG_REF",
            "DWGREF",
            "DWGREFNO",
            "DWGREFNUMBER"
        }, StringComparer.OrdinalIgnoreCase);

        private static string N(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            s = s.Trim();
            while (s.Contains("  ")) s = s.Replace("  ", " ");
            // If needed to reduce false negatives:
            // s = Regex.Replace(s, @"[.,;:/\\\-()]+", "");
            return s;
        }

        private static string CompositeKeyMain(string owner, string desc, string loc, string dwg)
        {
            return string.Join("|", N(owner), N(desc), N(loc), N(dwg));
        }

        private static string CompositeKeyPage(string owner, string desc)
        {
            return string.Join("|", N(owner), N(desc));
        }

        [CommandMethod("XING_MATCH_TABLE")]
        public void MatchTable()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            var ed = doc.Editor;
            try
            {
                var options = new PromptEntityOptions("\nSelect a crossing table (Main/Page):")
                {
                    AllowNone = false
                };
                options.SetRejectMessage("\nEntity must be a table.");
                options.AddAllowedClass(typeof(Table), true);

                var selection = ed.GetEntity(options);
                if (selection.Status != PromptStatus.OK)
                {
                    Log(ed, "MATCH TABLE cancelled.");
                    return;
                }

                using (doc.LockDocument())
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    var table = tr.GetObject(selection.ObjectId, OpenMode.ForRead) as Table;
                    if (table == null)
                    {
                        Log(ed, "Selected entity is not a table.");
                        return;
                    }

                    var tableType = DetectTableType(table);
                    if (tableType == TableSync.XingTableType.Unknown)
                    {
                        Log(ed, "Selected table type could not be determined. Command aborted.");
                        return;
                    }

                    if (tableType == TableSync.XingTableType.LatLong)
                    {
                        Log(ed, "Lat/Long tables are not supported by MATCH TABLE.");
                        return;
                    }

                    Log(ed, $"Table type detected: {tableType}.");

                    HashSet<string> duplicateKeys;
                    Dictionary<string, CrossingRecord> byKey, byComposite;
                    Dictionary<string, string> byCompositeXKey;
                    BuildIndexesFromTable(
                        table,
                        tableType,
                        out byKey,
                        out byComposite,
                        out byCompositeXKey,
                        out duplicateKeys,
                        msg => Log(ed, msg));

                    Log(ed, $"Indexed table: byKey={byKey.Count}, byComposite={byComposite.Count}");
                    if (byKey.Count == 0 && byComposite.Count == 0)
                    {
                        Log(ed, "Selected table did not provide any usable crossing rows.");
                    }

                    var db = doc.Database;
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    // ---------------------------------------------------------------------
                    // Collect extents of every table in the drawing.  Blocks whose
                    // insertion point lies within one of these extents are considered
                    // table‑embedded and should be ignored when matching/updating.
                    // ---------------------------------------------------------------------
                    var tableExtents = new List<Extents3d>();
                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (!btr.IsLayout) continue;

                        foreach (ObjectId id in btr)
                        {
                            var tblRef = tr.GetObject(id, OpenMode.ForRead) as Table;
                            if (tblRef == null) continue;

                            try
                            {
                                var ext = tblRef.GeometricExtents;
                                tableExtents.Add(ext);
                            }
                            catch
                            {
                                // ignore tables without extents
                            }
                        }
                    }

                    int totalXing2 = 0;
                    int matched = 0;
                    int updated = 0;
                    int skippedNoKey = 0;
                    int skippedNoMatch = 0;
                    int matchedByComposite = 0;
                    int errors = 0;

                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        // Only process modelspace/paperspace, skip block definition records
                        if (!btr.IsLayout) continue;

                        foreach (ObjectId entId in btr)
                        {
                            var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                            if (br == null) continue;

                            // Skip blocks whose insertion point lies within any table extents
                            var pos = br.Position;
                            bool insideTable = false;
                            foreach (var ext in tableExtents)
                            {
                                if (pos.X >= ext.MinPoint.X && pos.X <= ext.MaxPoint.X &&
                                    pos.Y >= ext.MinPoint.Y && pos.Y <= ext.MaxPoint.Y)
                                {
                                    insideTable = true;
                                    break;
                                }
                            }
                            if (insideTable)
                                continue;

                            var name = GetEffectiveBlockName(br, tr);
                            if (!string.Equals(name, "XING2", StringComparison.OrdinalIgnoreCase))
                                continue;

                            totalXing2++;

                            try
                            {
                                ProcessBlock(
                                    ed,
                                    br,
                                    tr,
                                    tableType,
                                    byKey,
                                    byComposite,
                                    byCompositeXKey,
                                    ref matched,
                                    ref updated,
                                    ref skippedNoKey,
                                    ref skippedNoMatch,
                                    ref matchedByComposite);
                            }
                            catch (System.Exception ex)
                            {
                                errors++;
                                Log(ed, $"Block {br.Handle}: {ex.Message}");
                            }
                        }
                    }

                    tr.Commit();

                    if (duplicateKeys.Count > 0)
                    {
                        var list = duplicateKeys
                            .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
                            .ToList();
                        Log(ed, $"Duplicate keys ignored: {string.Join(", ", list)}.");
                    }

                    Log(
                        ed,
                        $"MATCH TABLE summary -> XING2 blocks: {totalXing2}, matched: {matched}, matched(composite): {matchedByComposite}, updated: {updated}, skipped(no key): {skippedNoKey}, skipped(no match): {skippedNoMatch}, errors: {errors}.");
                }
            }
            catch (System.Exception ex)
            {
                Log(ed, $"MATCH TABLE failed: {ex.Message}");
            }
        }

        private static void ProcessBlock(
            Editor ed,
            BlockReference br,
            Transaction tr,
            TableSync.XingTableType tableType,
            IDictionary<string, CrossingRecord> byKey,
            IDictionary<string, CrossingRecord> byComposite,
            IDictionary<string, string> byCompositeXKey,
            ref int matched,
            ref int updated,
            ref int skippedNoKey,
            ref int skippedNoMatch,
            ref int matchedByComposite)
        {
            if (br == null) return;

            var handle = br.Handle.ToString();
            var keyValue = GetAttributeText(br, tr, CrossingAttributeTags);
            var normalizedKey = TableSync.NormalizeKeyForLookup(keyValue);

            CrossingRecord record = null;
            var matchedViaComposite = false;
            string compositeUsed = null;

            if (!string.IsNullOrEmpty(normalizedKey))
            {
                byKey?.TryGetValue(normalizedKey, out record);
            }

            if (record == null)
            {
                var bOwner = GetAttributeText(br, tr, OwnerAttributeTags);
                var bDesc = GetAttributeText(br, tr, DescriptionAttributeTags);

                if (tableType == TableSync.XingTableType.Main)
                {
                    var bLoc = GetAttributeText(br, tr, LocationAttributeTags);
                    var bDwg = GetAttributeText(br, tr, DwgRefAttributeTags);
                    compositeUsed = CompositeKeyMain(bOwner, bDesc, bLoc, bDwg);
                }
                else
                {
                    compositeUsed = CompositeKeyPage(bOwner, bDesc);
                }

                if (!string.IsNullOrWhiteSpace(compositeUsed) &&
                    byComposite != null &&
                    byComposite.TryGetValue(compositeUsed, out var rowRec))
                {
                    record = rowRec;
                    matchedViaComposite = true;
                    matchedByComposite++;
                    Log(ed, $"Block {handle}: matched by composite.");
                }
            }

            if (record == null)
            {
                if (string.IsNullOrEmpty(normalizedKey))
                {
                    skippedNoKey++;
                    Log(ed, $"Block {handle}: missing or invalid crossing key.");
                }
                else
                {
                    skippedNoMatch++;
                    Log(ed, $"Block {handle}: key '{keyValue}' not found in the selected table.");
                }
                return;
            }

            matched++;
            br.UpgradeOpen();

            var changed = false;

            // Decide what CROSSING value to write:
            // 1) Prefer the table's readable X (raw) if we got it.
            // 2) Else if we matched by composite and we have a normalized X from byCompositeXKey, use that.
            // 3) Else fall back to the block's current normalizedKey (no harm if same).
            string crossingFromTableRaw = record?.Crossing; // raw value read from the table (e.g., "X1")
            string crossingFromCompositeIndex = null;
            if (matchedViaComposite &&
                byCompositeXKey != null &&
                !string.IsNullOrEmpty(compositeUsed))
            {
                byCompositeXKey.TryGetValue(compositeUsed, out crossingFromCompositeIndex); // normalized form like "X1"
            }

            var crossingToWrite = !string.IsNullOrEmpty(crossingFromTableRaw)
                ? crossingFromTableRaw
                : (!string.IsNullOrEmpty(crossingFromCompositeIndex)
                    ? crossingFromCompositeIndex
                    : normalizedKey);

            if (!string.IsNullOrEmpty(crossingToWrite))
            {
                changed |= SetAttributeIfExists(br, tr, CrossingAttributeTags, crossingToWrite, null);
            }

            changed |= SetAttributeIfExists(br, tr, OwnerAttributeTags, record.Owner, null);
            changed |= SetAttributeIfExists(br, tr, DescriptionAttributeTags, record.Description, null);

            if (tableType == TableSync.XingTableType.Main)
            {
                changed |= SetAttributeIfExists(br, tr, LocationAttributeTags, record.Location, null);
                changed |= SetAttributeIfExists(br, tr, DwgRefAttributeTags, record.DwgRef, null);
            }

            if (changed)
            {
                updated++;
                br.RecordGraphicsModified(true);
                if (!string.IsNullOrEmpty(normalizedKey) && !matchedViaComposite)
                {
                    Log(ed, $"Block {handle}: updated using table key '{normalizedKey}'.");
                }
                else if (matchedViaComposite)
                {
                    Log(ed, $"Block {handle}: updated via composite match.");
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(normalizedKey) && !matchedViaComposite)
                {
                    Log(ed, $"Block {handle}: already aligned with table key '{normalizedKey}'.");
                }
                else if (matchedViaComposite)
                {
                    Log(ed, $"Block {handle}: already aligned via composite match.");
                }
            }
        }

        // Helpers below (unchanged from your original code except for ReadCellValue):

        private static string GetAttributeText(BlockReference br, Transaction tr, ISet<string> tags)
        {
            if (br == null || tags == null || br.AttributeCollection == null)
            {
                return string.Empty;
            }

            foreach (ObjectId attId in br.AttributeCollection)
            {
                if (!attId.IsValid) continue;
                var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (attRef == null) continue;
                if (tags.Contains(attRef.Tag))
                    return attRef.TextString;
            }

            return string.Empty;
        }

        private static bool SetAttributeIfExists(
            BlockReference br,
            Transaction tr,
            ISet<string> tags,
            string value,
            ISet<string> onlyIfMissing)
        {
            bool changed = false;
            if (br.AttributeCollection == null) return false;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (attRef == null) continue;

                if (!tags.Contains(attRef.Tag)) continue;

                // if we only set if missing and current is not empty, skip
                if (onlyIfMissing != null && onlyIfMissing.Contains(attRef.Tag) && !string.IsNullOrEmpty(attRef.TextString))
                    continue;

                if (!string.Equals(attRef.TextString, value ?? string.Empty))
                {
                    attRef.TextString = value ?? string.Empty;
                    changed = true;
                }
            }

            return changed;
        }

        private static string GetEffectiveBlockName(BlockReference br, Transaction tr)
        {
            if (br == null)
            {
                return string.Empty;
            }

            var btrId = br.DynamicBlockTableRecord != ObjectId.Null
                ? br.DynamicBlockTableRecord
                : br.BlockTableRecord;
            var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
            return btr.Name;
        }

        private static TableSync.XingTableType DetectTableType(Table table)
        {
            // detection logic unchanged – omitted for brevity
            return TableSync.XingTableType.Main; // placeholder; replace with original detection logic
        }

        private static void BuildIndexesFromTable(
            Table table,
            TableSync.XingTableType tableType,
            out Dictionary<string, CrossingRecord> byKey,
            out Dictionary<string, CrossingRecord> byComposite,
            out Dictionary<string, string> byCompositeXKey,
            out HashSet<string> duplicateKeys,
            Action<string> log)
        {
            // index building logic unchanged – omitted for brevity
            byKey = new Dictionary<string, CrossingRecord>();
            byComposite = new Dictionary<string, CrossingRecord>();
            byCompositeXKey = new Dictionary<string, string>();
            duplicateKeys = new HashSet<string>();
            // You can copy your existing implementation here.
        }

        // ReadCellValue now treats a single dash as an empty value
        private static string ReadCellValue(Table table, int row, int column)
        {
            if (table == null || row < 0 || column < 0)
            {
                return string.Empty;
            }

            if (row >= table.Rows.Count || column >= table.Columns.Count)
            {
                return string.Empty;
            }

            try
            {
                var cell = table.Cells[row, column];
                var text = cell?.TextString ?? string.Empty;
                var trimmed = (text ?? string.Empty).Trim();
                // Normalize single dash (often used as placeholder) to empty
                if (trimmed == "-" || trimmed == "—")
                {
                    return string.Empty;
                }
                return trimmed;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void Log(Editor ed, string message)
        {
            if (ed == null || string.IsNullOrEmpty(message))
            {
                return;
            }

            try
            {
                ed.WriteMessage("\n[CrossingManager] " + message);
            }
            catch
            {
                // Ignore logging failures.
            }
        }
    }
}
