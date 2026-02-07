using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using XingManager.Models;

namespace XingManager.Services
{
    /// <summary>
    /// Provides scan and persistence helpers for xing2 blocks.
    /// </summary>
    public class XingRepository
    {
        public const string BlockName = "xing2";
        public const string LatLongDictionaryKey = "XING2_LATLNG";

        private readonly Document _doc;

        public XingRepository(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        public Document Document => _doc;

        /// <summary>
        /// Scans the drawing for xing2 blocks and collects their attribute values.
        /// Block references inserted into table cells are ignored by checking their
        /// geometric extents against the bounding extents of each table in the drawing.
        /// </summary>
        public ScanResult ScanCrossings()
        {
            var db = _doc.Database;
            var ed = _doc.Editor;
            var records = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            var contexts = new Dictionary<ObjectId, DuplicateResolver.InstanceContext>();
            var latLongRows = new List<LatLongRowInfo>();

            Logger.Info(ed, "=== XingRepository.ScanCrossings START ===");

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Collect extents of all tables in the drawing.
                var tableExtents = new List<TableExtentInfo>();
                var crossingTableRowsByTableId = new Dictionary<ObjectId, Dictionary<string, CrossingTableRowData>>();
                var tableBubbleBlocks = new Dictionary<ObjectId, ObjectId>();
                var tableOwnerBtrByTableId = new Dictionary<ObjectId, ObjectId>();
                var tableHandleByTableId = new Dictionary<ObjectId, string>();
                var btab = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                int totalTablesFound = 0;

                foreach (ObjectId btrId in btab)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    if (!btr.IsLayout) continue;

                    foreach (ObjectId entId in btr)
                    {
                        var tbl = tr.GetObject(entId, OpenMode.ForRead) as Table;
                        if (tbl == null) continue;

                        totalTablesFound++;
                        int rows = 0, cols = 0;
                        try { rows = tbl.Rows.Count; cols = tbl.Columns.Count; } catch { }

                        Logger.Info(ed, $"SCAN: Found table {tbl.ObjectId.Handle}, rows={rows}, cols={cols}");

                        // Track which layout/paperspace this table belongs to (for duplicate resolver labels)
                        tableOwnerBtrByTableId[tbl.ObjectId] = btrId;
                        try { tableHandleByTableId[tbl.ObjectId] = tbl.Handle.ToString(); } catch { /* ignore */ }


                        try
                        {
                            var ext = tbl.GeometricExtents;
                            tableExtents.Add(new TableExtentInfo { TableId = tbl.ObjectId, Extents = ext });
                        }
                        catch
                        {
                            // ignore tables without extents
                        }

                        try
                        {
                            CollectLatLongRows(tbl, latLongRows);

                            var rowMap = CollectCrossingTableRowMap(tbl, ed);
                            Logger.Info(ed, $"SCAN: Table {tbl.ObjectId.Handle} returned {rowMap.Count} crossing rows");
                            if (rowMap.Count > 0)
                            {
                                crossingTableRowsByTableId[tbl.ObjectId] = rowMap;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            Logger.Info(ed, $"SCAN: Table {tbl.ObjectId.Handle} threw exception: {ex.Message}");
                        }
                    }
                }

                Logger.Info(ed, $"SCAN: Total tables found={totalTablesFound}, crossingTableRowsByTableId.Count={crossingTableRowsByTableId.Count}");

                // Build a map of layout BlockTableRecordId -> layout name.
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                var layoutNames = new Dictionary<ObjectId, string>();
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    layoutNames[layout.BlockTableRecordId] = layout.LayoutName;
                }

                var blockRefClass = RXClass.GetClass(typeof(BlockReference));

                foreach (ObjectId btrId in btab)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    if (!btr.IsLayout)
                        continue;

                    var spaceName = layoutNames.ContainsKey(btrId) ? layoutNames[btrId] : btr.Name;

                    foreach (ObjectId entId in btr)
                    {
                        if (!entId.ObjectClass.IsDerivedFrom(blockRefClass))
                            continue;

                        var br = (BlockReference)tr.GetObject(entId, OpenMode.ForRead);
                        if (br == null)
                            continue;

                        var blockEffectiveName = GetBlockName(br, tr);
                        if (!string.Equals(blockEffectiveName, BlockName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        // Determine if this block is inside a table or on a paper-space layout.
                        bool isInTable = false;
                        ObjectId tableId = ObjectId.Null;
                        try
                        {
                            var blkExt = br.GeometricExtents;
                            foreach (var te in tableExtents)
                            {
                                var tblExt = te.Extents;
                                bool xOverlaps = blkExt.MinPoint.X <= tblExt.MaxPoint.X && blkExt.MaxPoint.X >= tblExt.MinPoint.X;
                                bool yOverlaps = blkExt.MinPoint.Y <= tblExt.MaxPoint.Y && blkExt.MaxPoint.Y >= tblExt.MinPoint.Y;
                                if (xOverlaps && yOverlaps)
                                {
                                    isInTable = true;
                                    tableId = te.TableId;
                                    break;
                                }
                            }
                        }
                        catch
                        {
                            // ignore extents errors
                        }

                        // Capture attributes
                        var attributes = ReadAttributes(br, tr);
                        var crossing = SanitizeAttributeValue(GetValue(attributes, "CROSSING"));
                        if (string.IsNullOrEmpty(crossing))
                            continue;

                        var owner = SanitizeAttributeValue(GetValue(attributes, "OWNER"));
                        var description = SanitizeAttributeValue(GetValue(attributes, "DESCRIPTION"));
                        var location = SanitizeAttributeValue(GetValue(attributes, "LOCATION"));
                        var dwgRef = SanitizeAttributeValue(GetValue(attributes, "DWG_REF"));
                        string lat;
                        string lng;
                        string zone;
                        TryGetLatLong(br, tr, out lat, out lng, out zone);

                        var crossingKey = crossing.Trim().ToUpperInvariant();
                        CrossingRecord record;
                        if (!records.TryGetValue(crossingKey, out record))
                        {
                            record = new CrossingRecord
                            {
                                Crossing = crossing,
                                Owner = owner,
                                Description = description,
                                Location = location,
                                DwgRef = dwgRef,
                                Lat = lat,
                                Long = lng,
                                Zone = zone,
                                CanonicalInstance = ObjectId.Null
                            };
                            records.Add(crossingKey, record);
                        }

                        record.AllInstances.Add(entId);

                        // Create instance context; mark as ignored if in table or not in Model space
                        bool ignore = isInTable || !string.Equals(spaceName, "Model", StringComparison.OrdinalIgnoreCase);
                        contexts[entId] = new DuplicateResolver.InstanceContext
                        {
                            ObjectId = entId,
                            Crossing = crossing,
                            SpaceName = spaceName,
                            Owner = owner,
                            Description = description,
                            Location = location,
                            DwgRef = dwgRef,
                            Lat = lat,
                            Long = lng,
                            Zone = zone,
                            IgnoreForDuplicates = ignore,
                            IsTableInstance = isInTable
                        };

                        if (isInTable && tableId != ObjectId.Null && crossingTableRowsByTableId.ContainsKey(tableId))
                        {
                            tableBubbleBlocks[entId] = tableId;
                        }

                        // Prefer a model-space instance as canonical; otherwise choose first.
                        if (record.CanonicalInstance.IsNull &&
                            string.Equals(spaceName, "Model", StringComparison.OrdinalIgnoreCase))
                        {
                            record.CanonicalInstance = entId;
                            record.Crossing = crossing;
                            record.Owner = owner;
                            record.Description = description;
                            record.Location = location;
                            record.DwgRef = dwgRef;
                            record.Lat = lat;
                            record.Long = lng;
                            record.Zone = zone;
                        }
                        else if (record.CanonicalInstance.IsNull)
                        {
                            record.CanonicalInstance = entId;
                            record.Owner = owner;
                            record.Description = description;
                            record.Location = location;
                            record.DwgRef = dwgRef;
                            record.Lat = lat;
                            record.Long = lng;
                            record.Zone = zone;
                        }
                    }
                }

                // Apply table-row values into per-instance contexts so the duplicate resolver can
                // surface mismatches between drawing blocks and table rows.
                ApplyCrossingTableOverrides(tableBubbleBlocks, crossingTableRowsByTableId, records, contexts);

                // Attach crossing table row sources to the corresponding records so mismatches with
                // block/UI values can trigger the crossing duplicate resolver.
                ApplyCrossingTableRowsToRecords(crossingTableRowsByTableId, tableOwnerBtrByTableId, tableHandleByTableId, layoutNames, records);

                tr.Commit();
            }

            ApplyLatLongRows(latLongRows, records);

            var ordered = records.Values
                .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();

            Logger.Info(ed, $"=== XingRepository.ScanCrossings END: {ordered.Count} records ===");

            return new ScanResult
            {
                Records = ordered,
                InstanceContexts = contexts
            };
        }

        public void ApplyChanges(IList<CrossingRecord> records, TableSync tableSync)
        {
            if (records == null) throw new ArgumentNullException(nameof(records));
            if (tableSync == null) throw new ArgumentNullException(nameof(tableSync));

            var db = _doc.Database;

            using (_doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var record in records)
                {
                    foreach (var instanceId in record.AllInstances.Distinct())
                    {
                        if (!instanceId.IsValid)
                            continue;

                        var br = tr.GetObject(instanceId, OpenMode.ForWrite) as BlockReference;
                        if (br == null)
                            continue;

                        // Write block attributes from the in-memory record
                        WriteAttribute(tr, br, "CROSSING", record.Crossing);
                        WriteAttribute(tr, br, "OWNER", record.Owner);
                        WriteAttribute(tr, br, "DESCRIPTION", record.Description);
                        WriteAttribute(tr, br, "LOCATION", record.Location);
                        WriteAttribute(tr, br, "DWG_REF", record.DwgRef);
                        SetLatLong(br, tr, record.Lat, record.Long, record.Zone);
                    }
                }

                tr.Commit();
            }

            // Keep repository writes limited to block attributes.
            // Table synchronization is orchestrated by higher-level flows (e.g., ApplyToDrawing)
            // to preserve table-specific matching rules and update order.
        }

        public ObjectId InsertCrossing(CrossingRecord record, Point3d position)
        {
            if (record == null)
                throw new ArgumentNullException(nameof(record));

            var db = _doc.Database;

            using (_doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    if (!blockTable.Has(BlockName))
                        throw new InvalidOperationException("Block 'xing2' definition not found in this drawing—insert one or use INSERT to bring it in, then retry.");

                    var blockId = blockTable[BlockName];
                    var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    var blockDef = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);

                    var br = new BlockReference(position, blockId);
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    if (blockDef.HasAttributeDefinitions)
                    {
                        foreach (ObjectId attDefId in blockDef)
                        {
                            var attDef = tr.GetObject(attDefId, OpenMode.ForRead) as AttributeDefinition;
                            if (attDef == null || attDef.Constant)
                                continue;

                            var attRef = new AttributeReference();
                            attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                            attRef.TextString = attDef.TextString;
                            br.AttributeCollection.AppendAttribute(attRef);
                            tr.AddNewlyCreatedDBObject(attRef, true);
                        }
                    }

                    WriteAttribute(tr, br, "CROSSING", record.Crossing);
                    WriteAttribute(tr, br, "OWNER", record.Owner);
                    WriteAttribute(tr, br, "DESCRIPTION", record.Description);
                    WriteAttribute(tr, br, "LOCATION", record.Location);
                    WriteAttribute(tr, br, "DWG_REF", record.DwgRef);
                    SetLatLong(br, tr, record.Lat, record.Long, record.Zone);

                    tr.Commit();
                    return br.ObjectId;
                }
            }
        }

        public void DeleteInstances(IEnumerable<ObjectId> instanceIds)
        {
            if (instanceIds == null) return;

            using (_doc.LockDocument())
            using (var tr = _doc.Database.TransactionManager.StartTransaction())
            {
                foreach (var id in instanceIds.Distinct())
                {
                    if (!id.IsValid) continue;

                    var dbObject = tr.GetObject(id, OpenMode.ForWrite, false, true);
                    if (dbObject != null && !dbObject.IsErased)
                        dbObject.Erase(true);
                }

                tr.Commit();
            }
        }

        public void DeleteBlocksByCrossing(string crossing)
        {
            if (string.IsNullOrWhiteSpace(crossing))
                return;

            var targetKey = crossing.Trim().ToUpperInvariant();

            using (_doc.LockDocument())
            using (var tr = _doc.Database.TransactionManager.StartTransaction())
            {
                var db = _doc.Database;
                var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var blockRefClass = RXClass.GetClass(typeof(BlockReference));
                var idsToDelete = new List<ObjectId>();

                foreach (ObjectId btrId in blockTable)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    if (!btr.IsLayout)
                        continue;

                    foreach (ObjectId entId in btr)
                    {
                        if (!entId.ObjectClass.IsDerivedFrom(blockRefClass))
                            continue;

                        var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                        if (br == null || br.IsErased)
                            continue;

                        var blockName = GetBlockName(br, tr);
                        if (!string.Equals(blockName, BlockName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        var attributes = ReadAttributes(br, tr);
                        var attValue = GetValue(attributes, "CROSSING");
                        var attKey = (attValue ?? string.Empty).Trim().ToUpperInvariant();
                        if (string.IsNullOrEmpty(attKey))
                            continue;

                        if (string.Equals(attKey, targetKey, StringComparison.Ordinal))
                            idsToDelete.Add(entId);
                    }
                }

                foreach (var id in idsToDelete.Distinct())
                {
                    var br = tr.GetObject(id, OpenMode.ForWrite, false, true) as BlockReference;
                    if (br != null && !br.IsErased)
                        br.Erase(true);
                }

                tr.Commit();
            }
        }

        public bool TryGetLatLong(BlockReference br, Transaction tr, out string lat, out string lng, out string zone)
        {
            lat = string.Empty;
            lng = string.Empty;
            zone = string.Empty;
            if (br == null) return false;

            if (br.ExtensionDictionary.IsNull)
                return false;

            var dict = tr.GetObject(br.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
            if (dict == null || !dict.Contains(LatLongDictionaryKey))
                return false;

            var xrec = tr.GetObject(dict.GetAt(LatLongDictionaryKey), OpenMode.ForRead) as Xrecord;
            if (xrec == null || xrec.Data == null)
                return false;

            var values = xrec.Data.AsArray();
            if (values.Length >= 1)
                lat = Convert.ToString(values[0].Value, CultureInfo.InvariantCulture);

            if (values.Length >= 2)
                lng = Convert.ToString(values[1].Value, CultureInfo.InvariantCulture);

            if (values.Length >= 3)
                zone = Convert.ToString(values[2].Value, CultureInfo.InvariantCulture);

            return !string.IsNullOrEmpty(lat) || !string.IsNullOrEmpty(lng) || !string.IsNullOrEmpty(zone);
        }

        public void SetLatLong(BlockReference br, Transaction tr, string lat, string lng, string zone)
        {
            if (br == null) throw new ArgumentNullException(nameof(br));
            if (tr == null) throw new ArgumentNullException(nameof(tr));

            if (br.ExtensionDictionary.IsNull)
                br.CreateExtensionDictionary();

            var dict = (DBDictionary)tr.GetObject(br.ExtensionDictionary, OpenMode.ForWrite);
            Xrecord xrec;
            if (dict.Contains(LatLongDictionaryKey))
                xrec = (Xrecord)tr.GetObject(dict.GetAt(LatLongDictionaryKey), OpenMode.ForWrite);
            else
            {
                xrec = new Xrecord();
                dict.SetAt(LatLongDictionaryKey, xrec);
                tr.AddNewlyCreatedDBObject(xrec, true);
            }

            xrec.Data = new ResultBuffer(
                new TypedValue((int)DxfCode.Text, lat ?? string.Empty),
                new TypedValue((int)DxfCode.Text, lng ?? string.Empty),
                new TypedValue((int)DxfCode.Text, zone ?? string.Empty));
        }

        internal static void CollectLatLongRows(Table table, IList<LatLongRowInfo> rows)
        {
            if (table == null || rows == null)
                return;

            if (!IsLatLongTable(table))
                return;

            var startRow = TableSync.FindLatLongDataStartRow(table);
            if (startRow <= 0)
            {
                startRow = 0;
            }

            var columnCount = table.Columns.Count;
            var hasZoneColumn = columnCount >= 6;
            var hasDwgColumn = columnCount >= 6;

            var zoneColumn = hasZoneColumn ? 2 : -1;
            var latColumn = hasZoneColumn ? 3 : 2;
            var longColumn = hasZoneColumn ? 4 : 3;
            var dwgColumn = hasDwgColumn ? 5 : -1;

            var tableHandle = table.ObjectId.IsNull ? string.Empty : table.ObjectId.Handle.ToString();

            for (var row = startRow; row < table.Rows.Count; row++)
            {
                var crossing = TableSync.ResolveCrossingKey(table, row, 0);
                var description = TableSync.ReadCellTextSafe(table, row, 1);
                var zoneLabel = zoneColumn >= 0 ? TableSync.ReadCellTextSafe(table, row, zoneColumn) : string.Empty;
                var latitude = TableSync.ReadCellTextSafe(table, row, latColumn);
                var longitude = TableSync.ReadCellTextSafe(table, row, longColumn);
                var dwgRef = dwgColumn >= 0 ? TableSync.ReadCellTextSafe(table, row, dwgColumn) : string.Empty;

                if (string.IsNullOrWhiteSpace(crossing) &&
                    string.IsNullOrWhiteSpace(description) &&
                    string.IsNullOrWhiteSpace(latitude) &&
                    string.IsNullOrWhiteSpace(longitude) &&
                    string.IsNullOrWhiteSpace(zoneLabel) &&
                    string.IsNullOrWhiteSpace(dwgRef))
                {
                    continue;
                }

                rows.Add(new LatLongRowInfo
                {
                    TableId = table.ObjectId,
                    RowIndex = row,
                    SourceLabel = string.IsNullOrWhiteSpace(tableHandle)
                        ? string.Format(CultureInfo.InvariantCulture, "LAT/LONG table row {0}", row + 1)
                        : string.Format(CultureInfo.InvariantCulture, "LAT/LONG table {0} row {1}", tableHandle, row + 1),
                    Crossing = crossing,
                    Description = description,
                    Latitude = latitude,
                    Longitude = longitude,
                    Zone = ExtractZoneValue(zoneLabel),
                    DwgRef = dwgRef
                });
            }
        }

        internal static void ApplyLatLongRows(IEnumerable<LatLongRowInfo> rows, IDictionary<string, CrossingRecord> records)
        {
            if (rows == null || records == null)
                return;

            foreach (var record in records.Values)
            {
                record?.LatLongSources?.Clear();
            }

            foreach (var row in rows)
            {
                if (row == null)
                    continue;

                var record = FindRecordForLatLong(row, records);
                if (record == null)
                    continue;

                record.LatLongSources.Add(new CrossingRecord.LatLongSource
                {
                    SourceLabel = row.SourceLabel,
                    Description = row.Description,
                    Lat = row.Latitude,
                    Long = row.Longitude,
                    Zone = row.Zone,
                    DwgRef = row.DwgRef,
                    TableId = row.TableId,
                    RowIndex = row.RowIndex
                });

                var latitude = row.Latitude?.Trim();
                var longitude = row.Longitude?.Trim();
                var zone = row.Zone?.Trim();

                if (!string.IsNullOrWhiteSpace(latitude))
                {
                    record.Lat = latitude;
                }

                if (!string.IsNullOrWhiteSpace(longitude))
                {
                    record.Long = longitude;
                }

                if (!string.IsNullOrWhiteSpace(zone))
                {
                    record.Zone = zone;
                }

                var dwgRef = row.DwgRef?.Trim();
                if (!string.IsNullOrWhiteSpace(dwgRef) && string.IsNullOrWhiteSpace(record.DwgRef))
                {
                    record.DwgRef = dwgRef;
                }
            }

            foreach (var record in records.Values)
            {
                if (record == null || record.LatLongSources == null || record.LatLongSources.Count == 0)
                    continue;

                var normalizedLat = NormalizeForComparison(record.Lat);
                var normalizedLong = NormalizeForComparison(record.Long);
                var normalizedZone = NormalizeForComparison(record.Zone);

                var matching = record.LatLongSources.FirstOrDefault(s =>
                    string.Equals(NormalizeForComparison(s.Lat), normalizedLat, StringComparison.Ordinal) &&
                    string.Equals(NormalizeForComparison(s.Long), normalizedLong, StringComparison.Ordinal) &&
                    (string.IsNullOrEmpty(NormalizeForComparison(s.Zone)) ||
                     string.Equals(NormalizeForComparison(s.Zone), normalizedZone, StringComparison.Ordinal)));

                if (matching != null)
                    continue;

                var candidates = record.LatLongSources
                    .Where(s => !string.IsNullOrWhiteSpace(s.Lat) ||
                                !string.IsNullOrWhiteSpace(s.Long) ||
                                !string.IsNullOrWhiteSpace(s.Zone))
                    .ToList();

                if (candidates.Count == 0)
                    continue;

                var signatures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var candidate in candidates)
                {
                    var signature = string.Join("|",
                        NormalizeForComparison(candidate.Lat),
                        NormalizeForComparison(candidate.Long),
                        NormalizeForComparison(candidate.Zone));
                    signatures.Add(signature);
                }

                if (signatures.Count == 1)
                {
                    var source = candidates[0];
                    if (!string.IsNullOrWhiteSpace(source.Lat))
                        record.Lat = source.Lat.Trim();
                    if (!string.IsNullOrWhiteSpace(source.Long))
                        record.Long = source.Long.Trim();
                    if (!string.IsNullOrWhiteSpace(source.Zone))
                        record.Zone = source.Zone.Trim();
                }
            }
        }

        private static CrossingRecord FindRecordForLatLong(LatLongRowInfo row, IDictionary<string, CrossingRecord> records)
        {
            if (row == null || records == null)
                return null;

            var rawCrossing = row.Crossing ?? string.Empty;
            var normalizedCrossing = rawCrossing.Trim().ToUpperInvariant();
            var normalizedKey = TableSync.NormalizeKeyForLookup(rawCrossing);

            if (!string.IsNullOrWhiteSpace(normalizedCrossing) &&
                records.TryGetValue(normalizedCrossing, out var direct))
            {
                return direct;
            }

            if (!string.IsNullOrWhiteSpace(normalizedKey) &&
                records.TryGetValue(normalizedKey, out var normalizedMatch))
            {
                return normalizedMatch;
            }

            var byComparison = records.Values
                .FirstOrDefault(r => CrossingRecord.CompareCrossingKeys(r?.Crossing, rawCrossing) == 0);
            if (byComparison != null)
            {
                return byComparison;
            }

            var normalizedDescription = NormalizeForComparison(row.Description);
            if (string.IsNullOrEmpty(normalizedDescription))
            {
                return null;
            }

            var candidates = records.Values
                .Where(r => string.Equals(NormalizeForComparison(r?.Description), normalizedDescription, StringComparison.Ordinal))
                .ToList();

            if (candidates.Count == 1)
            {
                return candidates[0];
            }

            if (candidates.Count > 1)
            {
                var latNorm = NormalizeForComparison(row.Latitude);
                var longNorm = NormalizeForComparison(row.Longitude);
                var zoneNorm = NormalizeForComparison(row.Zone);
                var dwgNorm = NormalizeForComparison(row.DwgRef);

                var matches = candidates
                    .Where(r => string.Equals(NormalizeForComparison(r?.Lat), latNorm, StringComparison.Ordinal) &&
                                string.Equals(NormalizeForComparison(r?.Long), longNorm, StringComparison.Ordinal) &&
                                (string.IsNullOrEmpty(zoneNorm) || string.Equals(NormalizeForComparison(r?.Zone), zoneNorm, StringComparison.Ordinal)) &&
                                (string.IsNullOrEmpty(dwgNorm) || string.Equals(NormalizeForComparison(r?.DwgRef), dwgNorm, StringComparison.Ordinal)))
                    .ToList();

                if (matches.Count == 1)
                {
                    return matches[0];
                }
            }

            return null;
        }

        internal static bool IsLatLongTable(Table table)
        {
            if (table == null)
                return false;

            if (TableSync.HasMainHeaderRow(table) || TableSync.HasPageHeaderRow(table))
                return false;

            if ((table.Columns.Count != 4 && table.Columns.Count != 6) || table.Rows.Count <= 0)
                return false;

            if (TableSync.HasLatLongHeaderRow(table))
                return true;

            if (TableSync.FindLatLongDataStartRow(table) > 0)
                return true;

            var normalizedHeaders = new List<string>(4);
            var maxColumns = Math.Min(table.Columns.Count, 6);
            for (var column = 0; column < maxColumns; column++)
            {
                normalizedHeaders.Add(NormalizeForComparison(TableSync.ReadCellTextSafe(table, 0, column)));
            }

            var extendedHeaders = new[] { "ID", "DESCRIPTION", "ZONE", "LATITUDE", "LONGITUDE", "DWG_REF" };
            if (normalizedHeaders.Count == extendedHeaders.Length && normalizedHeaders.SequenceEqual(extendedHeaders))
                return true;

            var updatedHeaders = new[] { "ID", "DESCRIPTION", "LATITUDE", "LONGITUDE" };
            if (normalizedHeaders.Count == updatedHeaders.Length && normalizedHeaders.SequenceEqual(updatedHeaders))
                return true;

            var legacyHeaders = new[] { "XING", "DESCRIPTION", "LAT", "LONG" };
            if (normalizedHeaders.Count == legacyHeaders.Length && normalizedHeaders.SequenceEqual(legacyHeaders))
                return true;

            return LooksLikeLatLongDataRows(table);
        }

        private static string NormalizeForComparison(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            var normalized = TableSync.NormalizeText(value) ?? string.Empty;
            normalized = normalized.Trim();

            while (normalized.Contains("  "))
            {
                normalized = normalized.Replace("  ", " ");
            }

            return normalized.ToUpperInvariant();
        }

        private static bool LooksLikeLatLongDataRows(Table table)
        {
            if (table == null)
                return false;

            var columns = table.Columns.Count;
            if (columns < 4)
                return false;

            var latColumn = columns >= 6 ? 3 : 2;
            var longColumn = columns >= 6 ? 4 : 3;

            var rowCount = table.Rows.Count;
            if (rowCount <= 0)
                return false;

            var rowsToScan = Math.Min(rowCount, 12);
            var candidateRows = 0;

            for (var row = 0; row < rowsToScan; row++)
            {
                var latText = TableSync.ReadCellTextSafe(table, row, latColumn);
                var longText = TableSync.ReadCellTextSafe(table, row, longColumn);

                if (string.IsNullOrWhiteSpace(latText) && string.IsNullOrWhiteSpace(longText))
                    continue;

                var latIsCoordinate = LooksLikeCoordinate(latText, -90.0, 90.0);
                var longIsCoordinate = LooksLikeCoordinate(longText, -180.0, 180.0);

                if (!latIsCoordinate || !longIsCoordinate)
                {
                    if (LooksLikeHeader(latText) && LooksLikeHeader(longText))
                        continue;

                    return false;
                }

                var crossing = TableSync.ResolveCrossingKey(table, row, 0);
                var description = TableSync.ReadCellTextSafe(table, row, 1);

                if (!LooksLikeCrossingValue(crossing) && string.IsNullOrWhiteSpace(description))
                    return false;

                candidateRows++;
            }

            return candidateRows > 0;
        }

        private static bool LooksLikeCoordinate(string text, double min, double max)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var normalized = TableSync.NormalizeText(text);
            if (string.IsNullOrWhiteSpace(normalized))
                return false;

            if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            {
                if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
                    return false;
            }

            return value >= min && value <= max;
        }

        private static bool LooksLikeHeader(string text)
        {
            var normalized = NormalizeForComparison(text);
            if (string.IsNullOrWhiteSpace(normalized))
                return false;

            return normalized.StartsWith("LAT", StringComparison.Ordinal) ||
                   normalized.StartsWith("LONG", StringComparison.Ordinal) ||
                   normalized.Contains("ZONE");
        }

        private static bool LooksLikeCrossingValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            var normalized = NormalizeForComparison(value);
            if (normalized.StartsWith("X", StringComparison.Ordinal))
                return true;

            var token = CrossingRecord.ParseCrossingNumber(value);
            return token.Number > 0;
        }


        private struct TableExtentInfo
        {
            public ObjectId TableId;
            public Extents3d Extents;
        }

        private sealed class CrossingTableRowData
        {
            public string Owner = "";
            public string Description = "";
            public string Location = "";
            public string DwgRef = "";
            public bool HasOwner;
            public bool HasLocation;
            public bool HasDwgRef;
            public int RowIndex;
        }

        private static Dictionary<string, CrossingTableRowData> CollectCrossingTableRowMap(Table table, Editor ed)
        {
            var map = new Dictionary<string, CrossingTableRowData>(StringComparer.OrdinalIgnoreCase);

            if (table == null)
            {
                Logger.Info(ed, "  CollectCrossingTableRowMap: table is null");
                return map;
            }

            // Exclude LAT/LONG tables; those are handled by the lat/long resolver.
            if (IsLatLongTable(table))
            {
                Logger.Info(ed, "  CollectCrossingTableRowMap: Is LAT/LONG table, skipping");
                return map;
            }

            int rowCount;
            int colCount;
            try
            {
                rowCount = table.Rows.Count;
                colCount = table.Columns.Count;
                Logger.Info(ed, $"  CollectCrossingTableRowMap: rows={rowCount}, cols={colCount}");
            }
            catch (System.Exception ex)
            {
                Logger.Info(ed, $"  CollectCrossingTableRowMap: Failed to get row/col count: {ex.Message}");
                return map;
            }

            if (rowCount <= 0 || colCount < 2)
            {
                Logger.Info(ed, $"  CollectCrossingTableRowMap: Invalid dimensions (need rows>0, cols>=2), returning empty");
                return map;
            }

            // Supported crossing table shapes:
            //  - Main crossing table: 5 columns (XING / OWNER / DESCRIPTION / LOCATION / DWG_REF)
            //  - Crossing page table: 3 columns (XING / OWNER / DESCRIPTION) -> LOCATION & DWG_REF are ignored when comparing
            //  - Legacy (optional): 4 columns (XING / DESCRIPTION / LOCATION / DWG_REF) -> OWNER ignored when comparing
            bool hasOwner = false;
            bool hasLocation = false;
            bool hasDwgRef = false;
            int ownerCol = -1;
            int descCol = -1;
            int locCol = -1;
            int dwgCol = -1;

            if (colCount >= 5)
            {
                hasOwner = true; ownerCol = 1;
                descCol = 2;
                hasLocation = true; locCol = 3;
                hasDwgRef = true; dwgCol = 4;
                Logger.Info(ed, "  5+ col table: XING/OWNER/DESC/LOC/DWG");
            }
            else if (colCount == 3)
            {
                hasOwner = true; ownerCol = 1;
                descCol = 2;
                hasLocation = false;
                hasDwgRef = false;
                Logger.Info(ed, "  3-col table: XING/OWNER/DESC");
            }
            else if (colCount == 4)
            {
                // Legacy: no OWNER column
                hasOwner = false;
                descCol = 1;
                hasLocation = true; locCol = 2;
                hasDwgRef = true; dwgCol = 3;
                Logger.Info(ed, "  4-col table: XING/DESC/LOC/DWG (legacy)");
            }
            else
            {
                Logger.Info(ed, $"  CollectCrossingTableRowMap: Unsupported column count {colCount}, returning empty");
                return map;
            }

            int rowsAdded = 0;
            for (int r = 0; r < rowCount; r++)
            {
                string crossing = NormalizeCrossingKey(TableSync.ResolveCrossingKey(table, r, 0));
                Logger.Info(ed, $"    Row {r}: crossing='{crossing}'");

                if (!LooksLikeCrossingValue(crossing))
                {
                    Logger.Info(ed, $"      REJECTED by LooksLikeCrossingValue");
                    continue;
                }

                var rowData = new CrossingTableRowData
                {
                    RowIndex = r,
                    HasOwner = hasOwner,
                    HasLocation = hasLocation,
                    HasDwgRef = hasDwgRef
                };

                if (hasOwner && ownerCol >= 0 && ownerCol < colCount)
                    rowData.Owner = TableSync.ReadCellTextSafe(table, r, ownerCol) ?? string.Empty;

                if (descCol >= 0 && descCol < colCount)
                    rowData.Description = TableSync.ReadCellTextSafe(table, r, descCol) ?? string.Empty;

                if (hasLocation && locCol >= 0 && locCol < colCount)
                    rowData.Location = TableSync.ReadCellTextSafe(table, r, locCol) ?? string.Empty;

                if (hasDwgRef && dwgCol >= 0 && dwgCol < colCount)
                    rowData.DwgRef = TableSync.ReadCellTextSafe(table, r, dwgCol) ?? string.Empty;

                map[crossing] = rowData;
                rowsAdded++;
                Logger.Info(ed, $"      ADDED to map (desc='{rowData.Description}')");
            }

            Logger.Info(ed, $"  CollectCrossingTableRowMap: Returning {rowsAdded} crossing rows");
            return map;
        }

        private static void ApplyCrossingTableOverrides(
            Dictionary<ObjectId, ObjectId> tableBubbleBlocks,
            Dictionary<ObjectId, Dictionary<string, CrossingTableRowData>> crossingTableRowsByTableId,
            Dictionary<string, CrossingRecord> records,
            Dictionary<ObjectId, DuplicateResolver.InstanceContext> contexts)
        {
            if (tableBubbleBlocks == null || crossingTableRowsByTableId == null || records == null || contexts == null)
                return;

            foreach (var kvp in tableBubbleBlocks)
            {
                ObjectId blockId = kvp.Key;
                ObjectId tableId = kvp.Value;

                if (!contexts.TryGetValue(blockId, out var ctx) || ctx == null)
                    continue;

                string crossing = NormalizeCrossingKey(ctx.Crossing);
                if (string.IsNullOrWhiteSpace(crossing))
                    continue;

                if (!records.TryGetValue(crossing, out var rec) || rec == null)
                    continue;

                if (!crossingTableRowsByTableId.TryGetValue(tableId, out var rowMap) || rowMap == null)
                    continue;

                if (!rowMap.TryGetValue(crossing, out var rowData) || rowData == null)
                    continue;

                // For table instances, compare using the cell text (B-D), not the bubble's own attributes.
                ctx.Description = rowData.Description ?? string.Empty;
                ctx.Location = rowData.HasLocation ? (rowData.Location ?? string.Empty) : (rec.Location ?? string.Empty);
                ctx.DwgRef = rowData.HasDwgRef ? (rowData.DwgRef ?? string.Empty) : (rec.DwgRef ?? string.Empty);

                // Keep non-table fields aligned to the record so we don't create false duplicates.
                ctx.Owner = rowData.HasOwner ? (rowData.Owner ?? string.Empty) : (rec.Owner ?? string.Empty);
                ctx.Zone = rec.Zone ?? string.Empty;
                ctx.Lat = rec.Lat ?? string.Empty;
                ctx.Long = rec.Long ?? string.Empty;

                ctx.IgnoreForDuplicates = false;
            }
        }

        private static void ApplyCrossingTableRowsToRecords(
            IDictionary<ObjectId, Dictionary<string, CrossingTableRowData>> crossingTableRowsByTableId,
            IDictionary<ObjectId, ObjectId> tableOwnerBtrByTableId,
            IDictionary<ObjectId, string> tableHandleByTableId,
            IDictionary<ObjectId, string> layoutNames,
            IDictionary<string, CrossingRecord> records)
        {
            if (records == null)
                return;

            // Reset any previous scan's table sources
            foreach (var rec in records.Values)
            {
                rec?.CrossingTableSources?.Clear();
            }

            if (crossingTableRowsByTableId == null || crossingTableRowsByTableId.Count == 0)
                return;

            foreach (var tableEntry in crossingTableRowsByTableId)
            {
                var tableId = tableEntry.Key;
                var rowMap = tableEntry.Value;
                if (rowMap == null || rowMap.Count == 0)
                    continue;

                // Build a helpful label for the resolver (which layout + table handle)
                string layoutLabel = "Unknown";
                if (tableOwnerBtrByTableId != null &&
                    tableOwnerBtrByTableId.TryGetValue(tableId, out var ownerBtrId) &&
                    layoutNames != null &&
                    layoutNames.TryGetValue(ownerBtrId, out var ln) &&
                    !string.IsNullOrWhiteSpace(ln))
                {
                    layoutLabel = ln;
                }

                string handleLabel = "";
                if (tableHandleByTableId != null &&
                    tableHandleByTableId.TryGetValue(tableId, out var handle) &&
                    !string.IsNullOrWhiteSpace(handle))
                {
                    handleLabel = handle;
                }

                string sourceLabel = string.IsNullOrWhiteSpace(handleLabel)
                    ? $"TABLE ({layoutLabel})"
                    : $"TABLE ({layoutLabel}) #{handleLabel}";

                foreach (var rowEntry in rowMap)
                {
                    var crossingKey = NormalizeCrossingKey(rowEntry.Key);
                    if (string.IsNullOrWhiteSpace(crossingKey))
                        continue;

                    if (!records.TryGetValue(crossingKey, out var rec) || rec == null)
                        continue;

                    var row = rowEntry.Value;
                    if (row == null)
                        continue;

                    rec.CrossingTableSources.Add(new CrossingRecord.CrossingTableSource
                    {
                        SourceLabel = sourceLabel,
                        Owner = row.Owner ?? string.Empty,
                        Description = row.Description ?? string.Empty,
                        Location = row.Location ?? string.Empty,
                        DwgRef = row.DwgRef ?? string.Empty,
                        HasOwner = row.HasOwner,
                        HasLocation = row.HasLocation,
                        HasDwgRef = row.HasDwgRef,
                        TableId = tableId,
                        RowIndex = row.RowIndex
                    });
                }
            }
        }

        private static string NormalizeCrossingKey(string raw)
        {
            return string.IsNullOrWhiteSpace(raw) ? string.Empty : raw.Trim().ToUpperInvariant();
        }

        internal sealed class LatLongRowInfo
        {
            public ObjectId TableId { get; set; }
            public int RowIndex { get; set; }
            public string SourceLabel { get; set; }
            public string Crossing { get; set; }
            public string Description { get; set; }
            public string Latitude { get; set; }
            public string Longitude { get; set; }
            public string Zone { get; set; }
            public string DwgRef { get; set; }
        }

        private static string ExtractZoneValue(string zoneLabel)
        {
            if (string.IsNullOrWhiteSpace(zoneLabel))
            {
                return string.Empty;
            }

            var trimmed = TableSync.NormalizeText(zoneLabel) ?? string.Empty;
            trimmed = trimmed.Trim();
            if (trimmed.Length == 0)
            {
                return string.Empty;
            }

            var match = Regex.Match(trimmed, @"(\d+)");
            if (match.Success)
            {
                return match.Groups[1].Value.TrimStart('0');
            }

            if (trimmed.StartsWith("ZONE", StringComparison.OrdinalIgnoreCase))
            {
                var remainder = trimmed.Substring(4).Trim();
                return remainder.Length > 0 ? remainder : string.Empty;
            }

            return trimmed;
        }

        private static string GetBlockName(BlockReference br, Transaction tr)
        {
            if (br == null) return string.Empty;
            var btr = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
            return btr.Name;
        }

        private static Dictionary<string, string> ReadAttributes(BlockReference br, Transaction tr)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (br.AttributeCollection == null) return values;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                if (!attId.IsValid) continue;
                var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (attRef == null) continue;
                values[attRef.Tag] = attRef.TextString;
            }

            return values;
        }

        private static string GetValue(IDictionary<string, string> dict, string key)
        {
            return (dict != null && dict.TryGetValue(key, out var value)) ? value : string.Empty;
        }

        private static string SanitizeAttributeValue(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var sanitized = value;

            sanitized = Regex.Replace(sanitized, @"\\S([^;]+);", m =>
            {
                var frac = m.Groups[1].Value.Replace('#', '/');
                return frac;
            }, RegexOptions.IgnoreCase);

            sanitized = Regex.Replace(sanitized, @"\\P|\\~", " ", RegexOptions.IgnoreCase);
            sanitized = Regex.Replace(sanitized, @"\\[A-Za-z][^;]*;", string.Empty);
            sanitized = sanitized.Replace("{", string.Empty).Replace("}", string.Empty);
            sanitized = sanitized.Replace('\u00A0', ' ');
            sanitized = Regex.Replace(sanitized, @"\s+", " ").Trim();

            return sanitized;
        }

        private static void WriteAttribute(Transaction tr, BlockReference br, string tag, string value)
        {
            if (br.AttributeCollection == null) return;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (attRef == null) continue;

                if (string.Equals(attRef.Tag, tag, StringComparison.OrdinalIgnoreCase))
                    attRef.TextString = value ?? string.Empty;
            }
        }

        public class ScanResult
        {
            public IList<CrossingRecord> Records { get; set; }
            public IDictionary<ObjectId, DuplicateResolver.InstanceContext> InstanceContexts { get; set; }
        }
    }
}
