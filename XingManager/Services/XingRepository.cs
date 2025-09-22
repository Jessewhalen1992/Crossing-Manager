using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
                        TryGetLatLong(br, tr, out lat, out lng);

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
                        }
                    }
                }

                tr.Commit();
            }

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
            {
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

                            WriteAttribute(tr, br, "CROSSING", record.Crossing);
                            WriteAttribute(tr, br, "OWNER", record.Owner);
                            WriteAttribute(tr, br, "DESCRIPTION", record.Description);
                            WriteAttribute(tr, br, "LOCATION", record.Location);
                            WriteAttribute(tr, br, "DWG_REF", record.DwgRef);
                            SetLatLong(br, tr, record.Lat, record.Long);
                        }
                    }
                    tr.Commit();
                }

                tableSync.UpdateAllTables(_doc, records);
            }
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
                    SetLatLong(br, tr, record.Lat, record.Long);

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

        public bool TryGetLatLong(BlockReference br, Transaction tr, out string lat, out string lng)
        {
            lat = string.Empty;
            lng = string.Empty;
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

            return true;
        }

        public void SetLatLong(BlockReference br, Transaction tr, string lat, string lng)
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
                new TypedValue((int)DxfCode.Text, lng ?? string.Empty));
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
