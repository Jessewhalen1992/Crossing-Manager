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
    /// <summary>
    /// Reads a selected crossing table and pushes those values into XING2 blocks.
    /// Column A is a block cell; we read the CROSSING value from the cell's block attributes.
    /// Adjacent cells provide OWNER / DESCRIPTION / LOCATION / DWG_REF.
    /// </summary>
    public class TableMatcher
    {
        // --- Tag groups used when reading/writing block attributes ---
        private static readonly ISet<string> CrossingAttributeTags = new HashSet<string>(new[]
        {
            "CROSSING","XING","X_NO","XNUM","XNUMBER","NUMBER","INDEX","NO","LABEL"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> OwnerAttributeTags = new HashSet<string>(new[]
        {
            "OWNER","OWN","COMPANY"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> DescriptionAttributeTags = new HashSet<string>(new[]
        {
            "DESCRIPTION","DESC"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> LocationAttributeTags = new HashSet<string>(new[]
        {
            "LOCATION","LOC"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> DwgRefAttributeTags = new HashSet<string>(new[]
        {
            "DWG_REF","DWGREF","DWGREFNO","DWGREFNUMBER"
        }, StringComparer.OrdinalIgnoreCase);

        // --- simple normalizer for composite matching keys ---
        private static string N(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            s = s.Trim();
            while (s.Contains("  ")) s = s.Replace("  ", " ");
            return s;
        }
        private static string CompositeKeyMain(string owner, string desc, string loc, string dwg)
            => string.Join("|", N(owner), N(desc), N(loc), N(dwg));
        private static string CompositeKeyPage(string owner, string desc)
            => string.Join("|", N(owner), N(desc));

        [CommandMethod("XING_MATCH_TABLE")]
        public void MatchTable()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

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

                    // Determine table type with your existing logic.
                    var tableType = DetectTableType(table, tr);
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

                    // Build lookup dictionaries FROM the selected table (this is the "source of truth")
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

                    Log(ed, $"Indexed table rows -> byKey={byKey.Count}, byComposite={byComposite.Count}");

                    // Collect extents of all tables so we can ignore blocks embedded in tables
                    var db = doc.Database;
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var tableExtents = new List<Extents3d>();
                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (!btr.IsLayout) continue;

                        foreach (ObjectId id in btr)
                        {
                            var t = tr.GetObject(id, OpenMode.ForRead) as Table;
                            if (t == null) continue;
                            try { tableExtents.Add(t.GeometricExtents); } catch { }
                        }
                    }

                    int totalXing2 = 0, matched = 0, updated = 0, skippedNoKey = 0, skippedNoMatch = 0, matchedByComposite = 0, errors = 0;

                    // Update blocks in model & paper spaces (but skip those intersecting any table extents)
                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (!btr.IsLayout) continue;

                        foreach (ObjectId entId in btr)
                        {
                            var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                            if (br == null) continue;

                            bool insideTable = false;
                            try
                            {
                                var ext = br.GeometricExtents;
                                foreach (var tExt in tableExtents)
                                {
                                    bool xOver = ext.MinPoint.X <= tExt.MaxPoint.X && ext.MaxPoint.X >= tExt.MinPoint.X;
                                    bool yOver = ext.MinPoint.Y <= tExt.MaxPoint.Y && ext.MaxPoint.Y >= tExt.MinPoint.Y;
                                    if (xOver && yOver) { insideTable = true; break; }
                                }
                            }
                            catch { }
                            if (insideTable) continue;

                            var name = GetEffectiveBlockName(br, tr);
                            if (!string.Equals(name, "XING2", StringComparison.OrdinalIgnoreCase))
                                continue;

                            totalXing2++;
                            try
                            {
                                ProcessBlock(
                                    ed, br, tr, tableType,
                                    byKey, byComposite, byCompositeXKey,
                                    ref matched, ref updated, ref skippedNoKey, ref skippedNoMatch, ref matchedByComposite);
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
                        Log(ed, "Duplicate X keys in table ignored: " + string.Join(", ", duplicateKeys));

                    Log(ed, $"MATCH TABLE summary -> XING2: {totalXing2}, matched: {matched}, matched(composite): {matchedByComposite}, updated: {updated}, skipped(no key): {skippedNoKey}, skipped(no match): {skippedNoMatch}, errors: {errors}.");
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
            var normalizedKey = TableSync.NormalizeKeyForLookup(keyValue); // same normalizer as tables. :contentReference[oaicite:2]{index=2}

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
                    Log(ed, $"Block {handle}: missing/invalid CROSSING.");
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

            // CROSSING to write:
            // 1) prefer the raw X from the table; 2) else if matched via composite, use its X; 3) else keep block's.
            string crossingFromTableRaw = record.Crossing;
            string crossingFromCompositeIndex = null;
            if (matchedViaComposite && byCompositeXKey != null && !string.IsNullOrEmpty(compositeUsed))
                byCompositeXKey.TryGetValue(compositeUsed, out crossingFromCompositeIndex);

            var crossingToWrite = !string.IsNullOrEmpty(crossingFromTableRaw)
                ? crossingFromTableRaw
                : (!string.IsNullOrEmpty(crossingFromCompositeIndex) ? crossingFromCompositeIndex : normalizedKey);

            if (!string.IsNullOrEmpty(crossingToWrite))
                changed |= SetAttributeIfExists(br, tr, CrossingAttributeTags, crossingToWrite, null);

            // From table to block
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
                Log(ed, $"Block {handle}: attributes updated from table.");
            }
            else
            {
                Log(ed, $"Block {handle}: already aligned with table.");
            }
        }

        // ---------- Helpers used to build the table indexes ----------

        private static TableSync.XingTableType DetectTableType(Table table, Transaction tr)
        {
            try
            {
                // Reuse your existing IdentifyTable logic. :contentReference[oaicite:3]{index=3}
                var ts = new TableSync(new TableFactory());
                return ts.IdentifyTable(table, tr);
            }
            catch
            {
                return TableSync.XingTableType.Unknown;
            }
        }

        /// <summary>
        /// Reads each data row from the selected table and builds:
        ///  byKey           : "X#" (normalized) -> row record
        ///  byComposite     : composite(owner,desc[,loc,dwg]) -> row record
        ///  byCompositeXKey : composite -> "X#" (normalized)
        /// Tracks duplicate X#s encountered in duplicateKeys.
        /// </summary>
        private static void BuildIndexesFromTable(
            Table table,
            TableSync.XingTableType tableType,
            out Dictionary<string, CrossingRecord> byKey,
            out Dictionary<string, CrossingRecord> byComposite,
            out Dictionary<string, string> byCompositeXKey,
            out HashSet<string> duplicateKeys,
            Action<string> log)
        {
            byKey = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            byComposite = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            byCompositeXKey = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            duplicateKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (table == null) return;

            var rows = table.Rows.Count;
            var cols = table.Columns.Count;

            for (int row = 0; row < rows; row++)
            {
                // Column A: read CROSSING from the block cell (falls back to text). :contentReference[oaicite:4]{index=4}
                var rawKey = TableSync.ResolveCrossingKey(table, row, 0);
                var normalized = TableSync.NormalizeKeyForLookup(rawKey);
                if (string.IsNullOrEmpty(normalized))
                {
                    // Likely header/empty row; ignore.
                    continue;
                }

                // Adjacent cells
                var owner = (cols > 1) ? ReadCellValue(table, row, 1) : string.Empty;
                var desc = (cols > 2) ? ReadCellValue(table, row, 2) : string.Empty;

                string loc = string.Empty, dwg = string.Empty;
                if (tableType == TableSync.XingTableType.Main)
                {
                    loc = (cols > 3) ? ReadCellValue(table, row, 3) : string.Empty;
                    dwg = (cols > 4) ? ReadCellValue(table, row, 4) : string.Empty;
                }

                var rec = new CrossingRecord
                {
                    Crossing = (rawKey ?? string.Empty).Trim(),
                    Owner = owner,
                    Description = desc,
                    Location = loc,
                    DwgRef = dwg
                };

                if (byKey.ContainsKey(normalized))
                {
                    duplicateKeys.Add(normalized);
                }
                else
                {
                    byKey[normalized] = rec;
                }

                // Composite key (to match blocks even when X is off)
                var composite = (tableType == TableSync.XingTableType.Main)
                    ? CompositeKeyMain(owner, desc, loc, dwg)
                    : CompositeKeyPage(owner, desc);

                if (!string.IsNullOrWhiteSpace(composite) && !byComposite.ContainsKey(composite))
                {
                    byComposite[composite] = rec;
                    byCompositeXKey[composite] = normalized;
                }
            }

            if (duplicateKeys.Count > 0 && log != null)
                log("Table has duplicate X keys: " + string.Join(", ", duplicateKeys));
        }

        // ---------- Local cell/attribute helpers ----------

        private static string GetAttributeText(BlockReference br, Transaction tr, ISet<string> tags)
        {
            if (br == null || tags == null || br.AttributeCollection == null)
                return string.Empty;

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

        private static bool SetAttributeIfExists(BlockReference br, Transaction tr, ISet<string> tags, string value, ISet<string> onlyIfMissing)
        {
            bool changed = false;
            if (br.AttributeCollection == null) return false;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (attRef == null) continue;
                if (!tags.Contains(attRef.Tag)) continue;

                if (onlyIfMissing != null && onlyIfMissing.Contains(attRef.Tag) && !string.IsNullOrEmpty(attRef.TextString))
                    continue;

                var desired = value ?? string.Empty;
                if (!string.Equals(attRef.TextString, desired, StringComparison.Ordinal))
                {
                    attRef.TextString = desired;
                    changed = true;
                }
            }
            return changed;
        }

        private static string GetEffectiveBlockName(BlockReference br, Transaction tr)
        {
            if (br == null) return string.Empty;
            var btrId = br.DynamicBlockTableRecord != ObjectId.Null ? br.DynamicBlockTableRecord : br.BlockTableRecord;
            var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
            return btr.Name;
        }

        // Treat solitary "-" (and em-dash) as empty placeholders
        private static string ReadCellValue(Table table, int row, int column)
        {
            if (table == null || row < 0 || column < 0) return string.Empty;
            if (row >= table.Rows.Count || column >= table.Columns.Count) return string.Empty;

            try
            {
                var cell = table.Cells[row, column];
                var text = cell?.TextString ?? string.Empty;
                var trimmed = (text ?? string.Empty).Trim();
                if (trimmed == "-" || trimmed == "—") return string.Empty;
                return trimmed;
            }
            catch { return string.Empty; }
        }

        private static void Log(Editor ed, string message)
        {
            if (ed == null || string.IsNullOrEmpty(message)) return;
            try { ed.WriteMessage("\n[CrossingManager] " + message); } catch { }
        }
    }
}
