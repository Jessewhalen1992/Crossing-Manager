using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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

        private static readonly ISet<string> LatAttributeTags = new HashSet<string>(new[]
        {
            "LAT","LATITUDE"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> LongAttributeTags = new HashSet<string>(new[]
        {
            "LONG","LONGITUDE"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> ZoneAttributeTags = new HashSet<string>(new[]
        {
            "ZONE","ZONE_LABEL"
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
            using (Logger.Scope(ed, "match_table"))
            {
                try
                {
                    var options = new PromptEntityOptions("\nSelect a crossing table (Main/Page/LatLong):")
                    {
                        AllowNone = false
                    };
                    options.SetRejectMessage("\nEntity must be a table.");
                    options.AddAllowedClass(typeof(Table), true);

                    var selection = ed.GetEntity(options);
                    if (selection.Status != PromptStatus.OK)
                    {
                        Logger.Info(ed, "match_table status=cancelled");
                        return;
                    }

                using (doc.LockDocument())
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    var table = tr.GetObject(selection.ObjectId, OpenMode.ForRead) as Table;
                    if (table == null)
                    {
                        Logger.Warn(ed, "match_table status=aborted reason=not_table");
                        return;
                    }

                    // Determine table type with your existing logic.
                    var tableType = DetectTableType(table, tr);
                    if (tableType == TableSync.XingTableType.Unknown)
                    {
                        Logger.Warn(ed, "match_table status=aborted reason=unknown_table_type");
                        return;
                    }
                    Logger.Info(ed, $"table_detected type={tableType}");

                    // Build lookup dictionaries FROM the selected table (this is the "source of truth")
                    HashSet<string> duplicateKeys;
                    Dictionary<string, CrossingRecord> byKey, byComposite;
                    Dictionary<string, string> byCompositeXKey;
                    BuildIndexesFromTable(
                        table,
                        tableType,
                        ed,
                        out byKey,
                        out byComposite,
                        out byCompositeXKey,
                        out duplicateKeys);

                    Logger.Info(ed, $"indexed byKey={byKey.Count} byComposite={byComposite.Count} dupes={duplicateKeys.Count}");

                    // --------------------------------------------------------------------------------------
                    // When matching against a Main/Page table, enrich the in-memory records with LAT/LONG
                    // values from any LAT/LONG tables in the drawing **and** from existing XING blocks.
                    // Without this step, records created from Main/Page tables do not populate the
                    // Lat/Long fields and therefore will not update block extension dictionaries.
                    //
                    // First, harvest LAT/LONG values from all other LAT/LONG tables in the drawing.
                    // We walk through every table in every layout, identify those that are LAT/LONG tables,
                    // build temporary indexes from them, and merge their Lat/Long/Zone values into the
                    // primary byKey dictionary. Only missing fields are populated; if the selected table
                    // already contains LAT/LONG values, they are preserved.
                    if (tableType != TableSync.XingTableType.LatLong)
                    {
                        try
                        {
                            // Collect LAT/LONG information from other tables
                            var latLongByKey = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
                            var tableEnriched = 0;
                            var dbForLat = doc.Database;
                            var btable = (BlockTable)tr.GetObject(dbForLat.BlockTableId, OpenMode.ForRead);
                            foreach (ObjectId btrId2 in btable)
                            {
                                var btr2 = (BlockTableRecord)tr.GetObject(btrId2, OpenMode.ForRead);
                                if (!btr2.IsLayout) continue;
                                foreach (ObjectId entId2 in btr2)
                                {
                                    var tObj = tr.GetObject(entId2, OpenMode.ForRead) as Table;
                                    if (tObj == null) continue;
                                    // skip the selected table
                                    if (tObj.ObjectId == table.ObjectId) continue;

                                    TableSync.XingTableType tType = TableSync.XingTableType.Unknown;
                                    try
                                    {
                                        tType = DetectTableType(tObj, tr);
                                    }
                                    catch { }
                                    if (tType != TableSync.XingTableType.LatLong) continue;

                                    // Build indexes from this lat-long table
                                    BuildIndexesFromTable(
                                        tObj,
                                        tType,
                                        ed,
                                        out var llByKey,
                                        out var llByComposite,
                                        out var llByCompositeX,
                                        out var llDupes,
                                        logDuplicates: false);
                                    foreach (var kv in llByKey)
                                    {
                                        if (!latLongByKey.TryGetValue(kv.Key, out var merged))
                                        {
                                            latLongByKey[kv.Key] = kv.Value;
                                        }
                                        else
                                        {
                                            // combine lat/long/zone if missing
                                            if (string.IsNullOrWhiteSpace(merged.Lat) && !string.IsNullOrWhiteSpace(kv.Value.Lat))
                                                merged.Lat = kv.Value.Lat;
                                            if (string.IsNullOrWhiteSpace(merged.Long) && !string.IsNullOrWhiteSpace(kv.Value.Long))
                                                merged.Long = kv.Value.Long;
                                            if (string.IsNullOrWhiteSpace(merged.Zone) && !string.IsNullOrWhiteSpace(kv.Value.Zone))
                                                merged.Zone = kv.Value.Zone;
                                        }
                                    }
                                }
                            }
                            // Merge collected LAT/LONG values into byKey
                            foreach (var kv in latLongByKey)
                            {
                                if (byKey.TryGetValue(kv.Key, out var recTarget))
                                {
                                    var llRec = kv.Value;
                                    var changedFromTable = false;
                                    if (string.IsNullOrWhiteSpace(recTarget.Lat) && !string.IsNullOrWhiteSpace(llRec.Lat))
                                    {
                                        recTarget.Lat = llRec.Lat;
                                        changedFromTable = true;
                                    }
                                    if (string.IsNullOrWhiteSpace(recTarget.Long) && !string.IsNullOrWhiteSpace(llRec.Long))
                                    {
                                        recTarget.Long = llRec.Long;
                                        changedFromTable = true;
                                    }
                                    if (string.IsNullOrWhiteSpace(recTarget.Zone) && !string.IsNullOrWhiteSpace(llRec.Zone))
                                    {
                                        recTarget.Zone = llRec.Zone;
                                        changedFromTable = true;
                                    }
                                    if (changedFromTable)
                                        tableEnriched++;
                                }
                            }
                            Logger.Info(ed, $"enriched_from source=latlong_tables keys={tableEnriched}");
                        }
                        catch (System.Exception exTables)
                        {
                            // Best effort: do not abort if we cannot enrich from tables
                            Logger.Warn(ed, $"enrich_latlong_from_tables_failed err={exTables.Message}");
                        }

                        try
                        {
                            // Secondly, harvest LAT/LONG values from existing XING blocks via repository scan
                            var scanResult = new XingRepository(doc).ScanCrossings();
                            var blockEnriched = 0;
                            foreach (var record in scanResult.Records ?? new List<CrossingRecord>())
                            {
                                // Normalize the key using the same normalizer as BuildIndexesFromTable
                                var normalized = TableSync.NormalizeKeyForLookup(record.Crossing);
                                if (string.IsNullOrEmpty(normalized))
                                    continue;

                                if (byKey.TryGetValue(normalized, out var target))
                                {
                                    // Only populate missing fields; avoid overwriting values from the selected table
                                    var changedFromBlocks = false;
                                    if (string.IsNullOrWhiteSpace(target.Lat) && !string.IsNullOrWhiteSpace(record.Lat))
                                    {
                                        target.Lat = record.Lat;
                                        changedFromBlocks = true;
                                    }
                                    if (string.IsNullOrWhiteSpace(target.Long) && !string.IsNullOrWhiteSpace(record.Long))
                                    {
                                        target.Long = record.Long;
                                        changedFromBlocks = true;
                                    }
                                    if (string.IsNullOrWhiteSpace(target.Zone) && !string.IsNullOrWhiteSpace(record.Zone))
                                    {
                                        target.Zone = record.Zone;
                                        changedFromBlocks = true;
                                    }
                                    if (changedFromBlocks)
                                        blockEnriched++;
                                }
                            }
                            Logger.Info(ed, $"enriched_from source=blocks keys={blockEnriched}");
                        }
                        catch (System.Exception exScan)
                        {
                            // Best effort: do not abort the command on scan failures; just log the issue.
                            Logger.Warn(ed, $"enrich_latlong_from_blocks_failed err={exScan.Message}");
                        }
                    }

                    var repository = new XingRepository(doc);

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
                                    byKey, byComposite, byCompositeXKey, repository,
                                    ref matched, ref updated, ref skippedNoKey, ref skippedNoMatch, ref matchedByComposite);
                            }
                            catch (System.Exception ex)
                            {
                                errors++;
                                Logger.Error(ed, $"block handle={br.Handle} err={ex.Message}");
                            }
                        }
                    }

                    tr.Commit();

                    Logger.Info(ed, $"summary xing2_total={totalXing2} matched_key={matched} matched_composite={matchedByComposite} updated={updated} skipped_no_key={skippedNoKey} skipped_no_match={skippedNoMatch} errors={errors}");
                }
                }
                catch (System.Exception ex)
                {
                    Logger.Error(ed, $"match_table_failed err={ex.Message}");
                }
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
            XingRepository repository,
            ref int matched,
            ref int updated,
            ref int skippedNoKey,
            ref int skippedNoMatch,
            ref int matchedByComposite)
        {
            if (br == null) return;

            var handle = br.Handle.ToString();
            var keyValue = GetAttributeText(br, tr, CrossingAttributeTags);
            var normalizedKey = TableSync.NormalizeKeyForLookup(keyValue); // same normalizer as tables.

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
                    Logger.Debug(ed, $"block handle={handle} match=composite");
                }
            }

            if (record == null)
            {
                if (string.IsNullOrEmpty(normalizedKey))
                {
                    skippedNoKey++;
                    Logger.Info(ed, $"block handle={handle} skip reason=no_key");
                }
                else
                {
                    skippedNoMatch++;
                    Logger.Info(ed, $"block handle={handle} skip reason=no_match key={keyValue}");
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
            if (tableType == TableSync.XingTableType.LatLong)
            {
                if (!string.IsNullOrWhiteSpace(record.Description))
                    changed |= SetAttributeIfExists(br, tr, DescriptionAttributeTags, record.Description, null);

                // For LAT/LONG tables, do not write coordinates during MatchTable.
            }
            else
            {
                // MAIN or PAGE table: update textual attributes only
                changed |= SetAttributeIfExists(br, tr, OwnerAttributeTags, record.Owner, null);
                changed |= SetAttributeIfExists(br, tr, DescriptionAttributeTags, record.Description, null);

                if (tableType == TableSync.XingTableType.Main)
                {
                    changed |= SetAttributeIfExists(br, tr, LocationAttributeTags, record.Location, null);
                    changed |= SetAttributeIfExists(br, tr, DwgRefAttributeTags, record.DwgRef, null);
                }

                // DO NOT write LAT/LONG/ZONE values when matching Main or Page tables.
            }

            if (changed)
            {
                updated++;
                br.RecordGraphicsModified(true);
                Logger.Info(ed, $"block handle={handle} updated=true");
            }
            else
            {
                Logger.Debug(ed, $"block handle={handle} updated=false");
            }
        }

        // ---------- Helpers used to build the table indexes ----------
        private static TableSync.XingTableType DetectTableType(Table table, Transaction tr)
        {
            try
            {
                // Reuse your existing IdentifyTable logic.
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
        public static void BuildIndexesFromTable(
            Table table,
            TableSync.XingTableType tableType,
            Editor ed,
            out Dictionary<string, CrossingRecord> byKey,
            out Dictionary<string, CrossingRecord> byComposite,
            out Dictionary<string, string> byCompositeXKey,
            out HashSet<string> duplicateKeys,
            bool logDuplicates = true)
        {
            byKey = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            byComposite = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            byCompositeXKey = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            duplicateKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (table == null) return;

            var rows = table.Rows.Count;
            var cols = table.Columns.Count;

            var dataStartRow = 0;
            if (tableType == TableSync.XingTableType.LatLong)
            {
                var start = TableSync.FindLatLongDataStartRow(table);
                if (start > 0)
                    dataStartRow = start;
            }

            for (int row = dataStartRow; row < rows; row++)
            {
                // Column A: read CROSSING from the block cell (falls back to text).
                var rawKey = TableSync.ResolveCrossingKey(table, row, 0);
                var normalized = TableSync.NormalizeKeyForLookup(rawKey);
                if (string.IsNullOrEmpty(normalized))
                {
                    // Likely header/empty row; ignore.
                    continue;
                }

                // Adjacent cells
                string owner = string.Empty;
                string desc = string.Empty;
                string loc = string.Empty;
                string dwg = string.Empty;
                string lat = string.Empty;
                string lng = string.Empty;
                string zone = string.Empty;

                if (tableType == TableSync.XingTableType.LatLong)
                {
                    desc = (cols > 1) ? ReadCellValue(table, row, 1) : string.Empty;

                    var hasExtendedLayout = cols >= 5;
                    var zoneColumn = hasExtendedLayout ? 2 : -1;
                    var latColumn = hasExtendedLayout ? 3 : (cols > 2 ? 2 : -1);
                    var longColumn = hasExtendedLayout ? 4 : (cols > 3 ? 3 : -1);
                    var dwgColumn = cols >= 6 ? 5 : -1;

                    var zoneLabel = zoneColumn >= 0 ? ReadCellValue(table, row, zoneColumn) : string.Empty;
                    zone = ExtractZoneValue(zoneLabel);
                    lat = latColumn >= 0 ? ReadCellValue(table, row, latColumn) : string.Empty;
                    lng = longColumn >= 0 ? ReadCellValue(table, row, longColumn) : string.Empty;
                    dwg = dwgColumn >= 0 ? ReadCellValue(table, row, dwgColumn) : string.Empty;

                    if (string.IsNullOrWhiteSpace(desc) &&
                        string.IsNullOrWhiteSpace(lat) &&
                        string.IsNullOrWhiteSpace(lng) &&
                        string.IsNullOrWhiteSpace(zone) &&
                        string.IsNullOrWhiteSpace(dwg))
                    {
                        continue;
                    }
                }
                else
                {
                    owner = (cols > 1) ? ReadCellValue(table, row, 1) : string.Empty;
                    desc = (cols > 2) ? ReadCellValue(table, row, 2) : string.Empty;

                    if (tableType == TableSync.XingTableType.Main)
                    {
                        loc = (cols > 3) ? ReadCellValue(table, row, 3) : string.Empty;
                        dwg = (cols > 4) ? ReadCellValue(table, row, 4) : string.Empty;
                    }
                }

                var rec = new CrossingRecord
                {
                    Crossing = (rawKey ?? string.Empty).Trim(),
                    Owner = owner,
                    Description = desc,
                    Location = loc,
                    DwgRef = dwg,
                    Lat = lat,
                    Long = lng,
                    Zone = zone
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
                if (tableType != TableSync.XingTableType.LatLong)
                {
                    var composite = (tableType == TableSync.XingTableType.Main)
                        ? CompositeKeyMain(owner, desc, loc, dwg)
                        : CompositeKeyPage(owner, desc);

                    if (!string.IsNullOrWhiteSpace(composite) && !byComposite.ContainsKey(composite))
                    {
                        byComposite[composite] = rec;
                        byCompositeXKey[composite] = normalized;
                    }
                }
            }

            if (logDuplicates && duplicateKeys.Count > 0)
                Logger.Warn(ed, $"table_duplicate_keys count={duplicateKeys.Count} keys={string.Join(",", duplicateKeys)}");
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

        private static string ExtractZoneValue(string zoneLabel)
        {
            if (string.IsNullOrWhiteSpace(zoneLabel))
                return string.Empty;

            var normalized = TableSync.NormalizeText(zoneLabel) ?? string.Empty;
            normalized = normalized.Trim();
            if (normalized.Length == 0)
                return string.Empty;

            var match = Regex.Match(normalized, "(\\d+)");
            if (match.Success)
                return match.Groups[1].Value.TrimStart('0');

            if (normalized.StartsWith("ZONE", StringComparison.OrdinalIgnoreCase))
            {
                var remainder = normalized.Substring(4).Trim();
                return remainder.Length > 0 ? remainder : string.Empty;
            }

            return normalized;
        }

        private static bool ValuesEqual(string left, string right)
        {
            var leftNorm = (left ?? string.Empty).Trim();
            var rightNorm = (right ?? string.Empty).Trim();
            return string.Equals(leftNorm, rightNorm, StringComparison.OrdinalIgnoreCase);
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
                return (text ?? string.Empty).Trim();
            }
            catch { return string.Empty; }
        }

    }
}