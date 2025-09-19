using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
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
                    var byKey = BuildKeyIndexFromTable(table, tableType, out duplicateKeys, msg => Log(ed, msg));

                    if (byKey.Count == 0)
                    {
                        Log(ed, "Selected table did not provide any crossing rows (no readable X# in Column A).");
                    }

                    var db = doc.Database;
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    var totalXing2 = 0;
                    var matched = 0;
                    var updated = 0;
                    var skippedNoKey = 0;
                    var skippedNoMatch = 0;
                    var errors = 0;

                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        // Only process modelspace/paperspace, skip block definition records
                        if (!btr.IsLayout) continue;
                        foreach (ObjectId entId in btr)
                        {
                            var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                            if (br == null)
                            {
                                continue;
                            }

                            var name = GetEffectiveBlockName(br, tr);
                            if (!string.Equals(name, "XING2", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            totalXing2++;

                            try
                            {
                                ProcessBlock(ed, br, tr, tableType, byKey, ref matched, ref updated, ref skippedNoKey, ref skippedNoMatch);
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
                        var list = duplicateKeys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase).ToList();
                        Log(ed, $"Duplicate keys ignored: {string.Join(", ", list)}.");
                    }

                    Log(ed,
                        $"MATCH TABLE summary -> XING2 blocks: {totalXing2}, matched: {matched}, updated: {updated}, skipped(no key): {skippedNoKey}, skipped(no match): {skippedNoMatch}, errors: {errors}.");
                }
            }
            catch (System.Exception ex)
            {
                Log(ed, $"MATCH TABLE failed: {ex.Message}");
            }
        }

        private static void ProcessBlock(Editor ed, BlockReference br, Transaction tr, TableSync.XingTableType tableType,
            IDictionary<string, CrossingRecord> byKey,
            ref int matched, ref int updated, ref int skippedNoKey, ref int skippedNoMatch)
        {
            if (br == null)
            {
                return;
            }

            var handle = br.Handle.ToString();
            var keyValue = GetAttributeText(br, tr, CrossingAttributeTags);
            var normalizedKey = TableSync.NormalizeKeyForLookup(keyValue);

            if (string.IsNullOrEmpty(normalizedKey))
            {
                skippedNoKey++;
                Log(ed, $"Block {handle}: missing or invalid crossing key.");
                return;
            }

            if (byKey == null)
            {
                skippedNoMatch++;
                Log(ed, $"Block {handle}: key '{keyValue}' not found in the selected table.");
                return;
            }

            if (!byKey.TryGetValue(normalizedKey, out var record))
            {
                skippedNoMatch++;
                Log(ed, $"Block {handle}: key '{keyValue}' not found in the selected table.");
                return;
            }

            matched++;
            br.UpgradeOpen();

            var changed = false;
            changed |= SetAttributeIfExists(br, tr, CrossingAttributeTags, record.Crossing, null);
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
                Log(ed, $"Block {handle}: updated using table key '{normalizedKey}'.");
            }
            else
            {
                Log(ed, $"Block {handle}: already aligned with table key '{normalizedKey}'.");
            }
        }

        private static string GetEffectiveBlockName(BlockReference br, Transaction tr)
        {
            if (br == null)
            {
                return string.Empty;
            }

            var btrId = br.DynamicBlockTableRecord != ObjectId.Null ? br.DynamicBlockTableRecord : br.BlockTableRecord;
            var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
            return btr?.Name ?? string.Empty;
        }

        private static bool TryGetAttribute(BlockReference br, Transaction tr, ISet<string> tags, out AttributeReference attribute)
        {
            attribute = null;
            if (br == null || tr == null || tags == null || tags.Count == 0)
            {
                return false;
            }

            var collection = br.AttributeCollection;
            if (collection == null)
            {
                return false;
            }

            foreach (ObjectId attId in collection)
            {
                if (!attId.IsValid)
                {
                    continue;
                }

                var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (attRef == null)
                {
                    continue;
                }

                var tag = (attRef.Tag ?? string.Empty).Trim();
                if (tags.Contains(tag))
                {
                    attribute = attRef;
                    return true;
                }
            }

            return false;
        }

        private static string GetAttributeText(BlockReference br, Transaction tr, ISet<string> tags)
        {
            AttributeReference attribute;
            if (TryGetAttribute(br, tr, tags, out attribute))
            {
                return (attribute.TextString ?? string.Empty).Trim();
            }

            return string.Empty;
        }

        // Minimal detector for this command; we only support Main (5 cols) and Page (3 cols).
        private static TableSync.XingTableType DetectTableType(Table table)
        {
            if (table == null) return TableSync.XingTableType.Unknown;
            if (table.Columns.Count == 5) return TableSync.XingTableType.Main;
            if (table.Columns.Count == 3) return TableSync.XingTableType.Page;
            if (table.Columns.Count == 4) return TableSync.XingTableType.LatLong; // we will reject it explicitly
            return TableSync.XingTableType.Unknown;
        }

        private static bool SetAttributeIfExists(BlockReference br, Transaction tr, IEnumerable<string> tags, string value, Action<string> log)
        {
            if (br == null || tr == null || tags == null)
            {
                return false;
            }

            var tagSet = new HashSet<string>(tags, StringComparer.OrdinalIgnoreCase);
            if (tagSet.Count == 0)
            {
                return false;
            }

            var collection = br.AttributeCollection;
            if (collection == null)
            {
                return false;
            }

            var desired = value ?? string.Empty;
            var found = false;
            var updated = false;

            foreach (ObjectId attId in collection)
            {
                if (!attId.IsValid)
                {
                    continue;
                }

                var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (attRef == null)
                {
                    continue;
                }

                var tag = (attRef.Tag ?? string.Empty).Trim();
                if (!tagSet.Contains(tag))
                {
                    continue;
                }

                found = true;
                var existing = attRef.TextString ?? string.Empty;
                if (!string.Equals(existing, desired, StringComparison.Ordinal))
                {
                    attRef.UpgradeOpen();
                    attRef.TextString = desired;
                    updated = true;
                }
            }

            if (!found)
            {
                log?.Invoke($"Missing attribute tags ({string.Join(", ", tagSet)}) on block {br.Handle}.");
            }

            return updated;
        }

        private static Dictionary<string, CrossingRecord> BuildKeyIndexFromTable(
            Table table,
            TableSync.XingTableType type,
            out HashSet<string> duplicateKeys,
            Action<string> log)
        {
            var byKey = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            duplicateKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (table == null) return byKey;

            var rows = table.Rows.Count;
            for (int row = 0; row < rows; row++)
            {
                var rawKey = TableSync.ResolveCrossingKey(table, row, 0);
                var key = TableSync.NormalizeKeyForLookup(rawKey);
                if (string.IsNullOrEmpty(key)) continue;

                if (byKey.ContainsKey(key))
                {
                    duplicateKeys.Add(key);
                    log?.Invoke($"Row {row}: duplicate key '{key}' ignored; keeping the first occurrence.");
                    continue;
                }

                var rec = new CrossingRecord
                {
                    Crossing = (rawKey ?? string.Empty).Trim(),
                    Owner = ReadCellValue(table, row, 1),
                    Description = ReadCellValue(table, row, 2)
                };

                if (type == TableSync.XingTableType.Main)
                {
                    rec.Location = ReadCellValue(table, row, 3);
                    rec.DwgRef = ReadCellValue(table, row, 4);
                }

                byKey[key] = rec;
            }

            return byKey;
        }

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
                return text.Trim();
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
