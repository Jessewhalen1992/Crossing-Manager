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

        private static string N(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            s = s.Trim();
            while (s.Contains("  ")) s = s.Replace("  ", " ");
            // Optional (uncomment if you still miss matches due to punctuation):
            // s = Regex.Replace(s, @"[.,;:/\-()]+", "");
            return s;
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

                    Dictionary<string, CrossingRecord> byKey, byComposite;
                    Dictionary<string, string> byCompositeXKey;
                    HashSet<string> duplicateKeys;
                    BuildIndexesFromTable(table, tableType, out byKey, out byComposite, out byCompositeXKey, out duplicateKeys, msg => Log(ed, msg));

                    if (byKey.Count == 0 && byComposite.Count == 0)
                    {
                        Log(ed, "Selected table did not provide any usable crossing rows.");
                    }

                    var db = doc.Database;
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    var totalXing2 = 0;
                    var matched = 0;
                    var updated = 0;
                    var skippedNoKey = 0;
                    var skippedNoMatch = 0;
                    var matchedByComposite = 0;
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
                                ProcessBlock(ed, br, tr, tableType, byKey, byComposite, byCompositeXKey, ref matched, ref updated, ref skippedNoKey, ref skippedNoMatch, ref matchedByComposite);
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
                        $"MATCH TABLE summary -> XING2 blocks: {totalXing2}, matched: {matched}, updated: {updated}, skipped(no key): {skippedNoKey}, skipped(no match): {skippedNoMatch}, matched by composite: {matchedByComposite}, errors: {errors}.");
                }
            }
            catch (System.Exception ex)
            {
                Log(ed, $"MATCH TABLE failed: {ex.Message}");
            }
        }

        private static void ProcessBlock(Editor ed, BlockReference br, Transaction tr, TableSync.XingTableType tableType,
            IDictionary<string, CrossingRecord> byKey,
            IDictionary<string, CrossingRecord> byComposite,
            IDictionary<string, string> byCompositeXKey,
            ref int matched, ref int updated, ref int skippedNoKey, ref int skippedNoMatch, ref int matchedByComposite)
        {
            if (br == null)
            {
                return;
            }

            var handle = br.Handle.ToString();
            var keyValue = GetAttributeText(br, tr, CrossingAttributeTags);
            var normalizedKey = TableSync.NormalizeKeyForLookup(keyValue);
            var hadKey = !string.IsNullOrEmpty(normalizedKey);
            CrossingRecord record = null;
            var matchedViaComposite = false;
            string matchedCompositeKey = null;

            if (hadKey)
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
                    var comp = CompositeKeyMain(bOwner, bDesc, bLoc, bDwg);
                    matchedCompositeKey = comp;
                    CrossingRecord rowRec;
                    if (!string.IsNullOrWhiteSpace(comp) && byComposite != null && byComposite.TryGetValue(comp, out rowRec))
                    {
                        record = rowRec;
                        matchedViaComposite = true;
                        matchedByComposite++;
                        Log(ed, $"Block {handle}: matched by composite (owner/desc/location/dwg).");
                    }
                }
                else
                {
                    var comp = CompositeKeyPage(bOwner, bDesc);
                    matchedCompositeKey = comp;
                    CrossingRecord rowRec;
                    if (!string.IsNullOrWhiteSpace(comp) && byComposite != null && byComposite.TryGetValue(comp, out rowRec))
                    {
                        record = rowRec;
                        matchedViaComposite = true;
                        matchedByComposite++;
                        Log(ed, $"Block {handle}: matched by composite (owner/desc).");
                    }
                }
            }

            if (record == null)
            {
                if (hadKey)
                {
                    skippedNoMatch++;
                    Log(ed, $"Block {handle}: key '{keyValue}' not found in the selected table.");
                }
                else
                {
                    skippedNoKey++;
                    Log(ed, $"Block {handle}: missing or invalid crossing key.");
                }

                return;
            }

            matched++;

            string crossingToSet = null;

            if (!string.IsNullOrEmpty(normalizedKey)) crossingToSet = normalizedKey;

            if (crossingToSet == null && matchedViaComposite && !string.IsNullOrWhiteSpace(matchedCompositeKey))
            {
                string xFromTable;
                if (byCompositeXKey != null && byCompositeXKey.TryGetValue(matchedCompositeKey, out xFromTable) && !string.IsNullOrEmpty(xFromTable))
                    crossingToSet = xFromTable;
            }

            if (crossingToSet == null) crossingToSet = normalizedKey;

            br.UpgradeOpen();
            var changed = false;
            if (!string.IsNullOrEmpty(crossingToSet))
                changed |= SetAttributeIfExists(br, tr, CrossingAttributeTags, crossingToSet, null);
            changed |= SetAttributeIfExists(br, tr, OwnerAttributeTags, record.Owner, null);
            changed |= SetAttributeIfExists(br, tr, DescriptionAttributeTags, record.Description, null);

            if (tableType == TableSync.XingTableType.Main)
            {
                changed |= SetAttributeIfExists(br, tr, LocationAttributeTags, record.Location, null);
                changed |= SetAttributeIfExists(br, tr, DwgRefAttributeTags, record.DwgRef, null);
            }

            var logKey = !string.IsNullOrEmpty(crossingToSet)
                ? crossingToSet
                : (record?.Crossing ?? normalizedKey ?? string.Empty);

            if (changed)
            {
                updated++;
                br.RecordGraphicsModified(true);
                Log(ed, $"Block {handle}: updated using table key '{logKey}'.");
            }
            else
            {
                Log(ed, $"Block {handle}: already aligned with table key '{logKey}'.");
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

        private static Dictionary<string, CrossingRecord> BuildMapFromTable(Table table, TableSync.XingTableType type)
        {
            HashSet<string> duplicateKeys;
            return BuildMapFromTable(table, type, out duplicateKeys, null);
        }

        private static Dictionary<string, CrossingRecord> BuildMapFromTable(Table table, TableSync.XingTableType type, out HashSet<string> duplicateKeys, Action<string> log)
        {
            var map = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            duplicateKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (table == null)
            {
                return map;
            }

            var rowCount = table.Rows.Count;
            for (var row = 0; row < rowCount; row++)
            {
                var rawKey = TableSync.ResolveCrossingKey(table, row, 0);
                var normalized = TableSync.NormalizeKeyForLookup(rawKey);
                if (string.IsNullOrEmpty(normalized))
                {
                    continue;
                }

                if (map.ContainsKey(normalized))
                {
                    duplicateKeys.Add(normalized);
                    log?.Invoke($"Row {row}: duplicate key '{normalized}' ignored; keeping the first occurrence.");
                    continue;
                }

                var record = new CrossingRecord
                {
                    Crossing = (rawKey ?? string.Empty).Trim(),
                    Owner = ReadCellValue(table, row, 1),
                    Description = ReadCellValue(table, row, 2)
                };

                if (type == TableSync.XingTableType.Main)
                {
                    record.Location = ReadCellValue(table, row, 3);
                    record.DwgRef = ReadCellValue(table, row, 4);
                }

                map[normalized] = record;
            }

            return map;
        }

        private static string CompositeKeyMain(string owner, string desc, string loc, string dwg)
        {
            return string.Join("|", N(owner), N(desc), N(loc), N(dwg));
        }

        private static string CompositeKeyPage(string owner, string desc)
        {
            return string.Join("|", N(owner), N(desc));
        }

        private static void BuildIndexesFromTable(
            Table table,
            TableSync.XingTableType type,
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
            for (int row = 0; row < rows; row++)
            {
                var rawKey = TableSync.ResolveCrossingKey(table, row, 0);
                var key = TableSync.NormalizeKeyForLookup(rawKey);
                var owner = ReadCellValue(table, row, 1);
                var desc = ReadCellValue(table, row, 2);

                var rec = new CrossingRecord
                {
                    Crossing = (rawKey ?? string.Empty).Trim(),
                    Owner = owner,
                    Description = desc
                };

                string comp;
                if (type == TableSync.XingTableType.Main)
                {
                    var loc = ReadCellValue(table, row, 3);
                    var dwg = ReadCellValue(table, row, 4);
                    rec.Location = loc;
                    rec.DwgRef = dwg;
                    comp = CompositeKeyMain(owner, desc, loc, dwg);
                }
                else // Page
                {
                    comp = CompositeKeyPage(owner, desc);
                }

                if (!string.IsNullOrEmpty(key))
                {
                    if (byKey.ContainsKey(key))
                    {
                        duplicateKeys.Add(key);
                        log?.Invoke($"Row {row}: duplicate key '{key}' ignored; keeping the first occurrence.");
                    }
                    else
                    {
                        byKey[key] = rec;
                    }
                }

                if (!string.IsNullOrWhiteSpace(comp) && !byComposite.ContainsKey(comp))
                {
                    byComposite[comp] = rec;
                }

                if (!string.IsNullOrWhiteSpace(comp) && !byCompositeXKey.ContainsKey(comp))
                {
                    byCompositeXKey[comp] = key ?? string.Empty;
                }
            }
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
