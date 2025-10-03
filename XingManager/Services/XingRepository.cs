using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
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
            var records = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            var contexts = new Dictionary<ObjectId, DuplicateResolver.InstanceContext>();
            var latLongRows = new List<LatLongRowInfo>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Collect extents of all tables in the drawing.
                var tableExtents = new List<Extents3d>();
                var btab = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in btab)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    if (!btr.IsLayout) continue;

                    foreach (ObjectId entId in btr)
                    {
                        var tbl = tr.GetObject(entId, OpenMode.ForRead) as Table;
                        if (tbl == null) continue;

                        try
                        {
                            var ext = tbl.GeometricExtents;
                            tableExtents.Add(ext);
                        }
                        catch
                        {
                            // ignore tables without extents
                        }

                        try
                        {
                            CollectLatLongRows(tbl, latLongRows);
                        }
                        catch
                        {
                            // ignore lat/long parsing errors
                        }
                    }
                }

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
                        try
                        {
                            var blkExt = br.GeometricExtents;
                            foreach (var tblExt in tableExtents)
                            {
                                bool xOverlaps = blkExt.MinPoint.X <= tblExt.MaxPoint.X && blkExt.MaxPoint.X >= tblExt.MinPoint.X;
                                bool yOverlaps = blkExt.MinPoint.Y <= tblExt.MaxPoint.Y && blkExt.MaxPoint.Y >= tblExt.MinPoint.Y;
                                if (xOverlaps && yOverlaps)
                                {
                                    isInTable = true;
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
                        var crossing = GetValue(attributes, "CROSSING");
                        if (string.IsNullOrEmpty(crossing))
                            continue;

                        var owner = GetValue(attributes, "OWNER");
                        var description = GetValue(attributes, "DESCRIPTION");
                        var location = GetValue(attributes, "LOCATION");
                        var dwgRef = GetValue(attributes, "DWG_REF");
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
                            IgnoreForDuplicates = ignore
                        };

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

                tr.Commit();
            }

            ApplyLatLongRows(latLongRows, records);

            var ordered = records.Values
                .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();

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

            // IMPORTANT: removed automatic table updates here.
            // Tables will ONLY change when you run XING_MATCH_TABLE (table -> blocks)
            // or when you explicitly execute a table creation/update command.
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

        private static void CollectLatLongRows(Table table, IList<LatLongRowInfo> rows)
        {
            if (table == null || rows == null)
                return;

            if (!IsLatLongTable(table))
                return;

            var startRow = TableSync.FindLatLongDataStartRow(table);
            if (startRow <= 0)
            {
                startRow = 1;
            }

            var columnCount = table.Columns.Count;
            var hasZoneColumn = columnCount >= 6;
            var hasDwgColumn = columnCount >= 6;

            var zoneColumn = hasZoneColumn ? 2 : -1;
            var latColumn = hasZoneColumn ? 3 : 2;
            var longColumn = hasZoneColumn ? 4 : 3;
            var dwgColumn = hasDwgColumn ? 5 : -1;

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
                    Crossing = crossing,
                    Description = description,
                    Latitude = latitude,
                    Longitude = longitude,
                    Zone = ExtractZoneValue(zoneLabel),
                    DwgRef = dwgRef
                });
            }
        }

        private static void ApplyLatLongRows(IEnumerable<LatLongRowInfo> rows, IDictionary<string, CrossingRecord> records)
        {
            if (rows == null || records == null)
                return;

            foreach (var row in rows)
            {
                if (row == null)
                    continue;

                var record = FindRecordForLatLong(row, records);
                if (record == null)
                    continue;

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

        private static bool IsLatLongTable(Table table)
        {
            if (table == null)
                return false;

            if ((table.Columns.Count != 4 && table.Columns.Count != 6) || table.Rows.Count <= 0)
                return false;

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
            return normalizedHeaders.Count == legacyHeaders.Length && normalizedHeaders.SequenceEqual(legacyHeaders);
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

        private sealed class LatLongRowInfo
        {
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
