using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using XingManager.Models;

namespace XingManager.Services
{
    /// <summary>
    /// Synchronises AutoCAD tables with crossing data.
    /// </summary>
    public class TableSync
    {
        public enum XingTableType
        {
            Unknown,
            Main,
            Page,
            LatLong
        }

        private readonly TableFactory _factory;

        public TableSync(TableFactory factory)
        {
            _factory = factory ?? throw new ArgumentNullException("factory");
        }

        public void UpdateAllTables(Document doc, IList<CrossingRecord> records)
        {
            if (doc == null)
            {
                throw new ArgumentNullException("doc");
            }

            if (records == null)
            {
                throw new ArgumentNullException("records");
            }

            var db = doc.Database;
            var editor = doc.Editor;
            var byKey = records.ToDictionary(r => r.CrossingKey, r => r, StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var tableClass = RXClass.GetClass(typeof(Table));

                foreach (ObjectId btrId in blockTable)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    foreach (ObjectId entId in btr)
                    {
                        if (!entId.ObjectClass.IsDerivedFrom(tableClass))
                        {
                            continue;
                        }

                        var table = (Table)tr.GetObject(entId, OpenMode.ForRead);
                        var type = IdentifyTable(table, tr);
                        if (type == XingTableType.Unknown)
                        {
                            continue;
                        }

                        table.UpgradeOpen();
                        try
                        {
                            switch (type)
                            {
                                case XingTableType.Main:
                                    UpdateMainTable(table, byKey);
                                    break;
                                case XingTableType.Page:
                                    UpdatePageTable(table, byKey);
                                    break;
                                case XingTableType.LatLong:
                                    UpdateLatLongTable(table, byKey);
                                    break;
                            }

                            _factory.TagTable(tr, table, type.ToString().ToUpperInvariant());
                        }
                        catch (System.Exception ex)
                        {
                            editor.WriteMessage("\n[CrossingManager] Failed to update table {0}: {1}", entId.Handle, ex.Message);
                        }
                    }
                }

                tr.Commit();
            }
        }

        public Table CreateAndInsertPageTable(Database db, Transaction tr, BlockTableRecord space, Point3d insertPoint, string dwgRef, IEnumerable<CrossingRecord> records)
        {
            var table = _factory.CreateCrossingPageTable(db, tr, dwgRef, records);
            table.Position = insertPoint;
            space.UpgradeOpen();
            space.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
            _factory.TagTable(tr, table, XingTableType.Page.ToString().ToUpperInvariant());
            return table;
        }

        public Table CreateAndInsertLatLongTable(Database db, Transaction tr, BlockTableRecord space, Point3d insertPoint, CrossingRecord record)
        {
            var table = _factory.CreateLatLongTable(db, tr, record);
            table.Position = insertPoint;
            space.UpgradeOpen();
            space.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
            _factory.TagTable(tr, table, XingTableType.LatLong.ToString().ToUpperInvariant());
            return table;
        }

        public XingTableType IdentifyTable(Table table, Transaction tr)
        {
            if (table == null)
            {
                return XingTableType.Unknown;
            }

            if (!table.ExtensionDictionary.IsNull)
            {
                var dict = (DBDictionary)tr.GetObject(table.ExtensionDictionary, OpenMode.ForRead);
                if (dict.Contains(TableFactory.TableTypeXrecordKey))
                {
                    var xrec = (Xrecord)tr.GetObject(dict.GetAt(TableFactory.TableTypeXrecordKey), OpenMode.ForRead);
                    var data = xrec.Data;
                    if (data != null)
                    {
                        foreach (TypedValue value in data)
                        {
                            if (value.TypeCode == (int)DxfCode.Text)
                            {
                                var text = Convert.ToString(value.Value, CultureInfo.InvariantCulture);
                                if (string.Equals(text, "MAIN", StringComparison.OrdinalIgnoreCase))
                                {
                                    return XingTableType.Main;
                                }
                                if (string.Equals(text, "PAGE", StringComparison.OrdinalIgnoreCase))
                                {
                                    return XingTableType.Page;
                                }
                                if (string.Equals(text, "LATLONG", StringComparison.OrdinalIgnoreCase))
                                {
                                    return XingTableType.LatLong;
                                }
                            }
                        }
                    }
                }
            }

            if (table.Columns.Count == 5 && table.Rows.Count >= 1)
            {
                var headers = ReadHeaders(table, 5);
                if (IsMainHeader(headers))
                {
                    return XingTableType.Main;
                }
            }

            if (table.Columns.Count == 3 && table.Rows.Count >= 1)
            {
                var headers = ReadHeaders(table, 3);
                if (IsPageHeader(headers))
                {
                    return XingTableType.Page;
                }
            }

            if (table.Columns.Count == 4 && table.Rows.Count >= 1)
            {
                var headers = ReadHeaders(table, 4);
                if (IsLatLongHeader(headers))
                {
                    return XingTableType.LatLong;
                }
            }

            return XingTableType.Unknown;
        }

        private static bool IsMainHeader(List<string> headers)
        {
            var expected = new[] { "XING", "OWNER", "DESCRIPTION", "LOCATION", "DWG_REF" };
            return CompareHeaders(headers, expected);
        }

        private static bool IsPageHeader(List<string> headers)
        {
            var expected = new[] { "XING", "OWNER", "DESCRIPTION" };
            return CompareHeaders(headers, expected);
        }

        private static bool IsLatLongHeader(List<string> headers)
        {
            var expected = new[] { "XING", "DESCRIPTION", "LAT", "LONG" };
            return CompareHeaders(headers, expected);
        }

        private static bool CompareHeaders(List<string> headers, string[] expected)
        {
            if (headers == null || headers.Count != expected.Length)
            {
                return false;
            }

            for (var i = 0; i < expected.Length; i++)
            {
                if (!string.Equals(headers[i], expected[i], StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }

        private static List<string> ReadHeaders(Table table, int columns)
        {
            var list = new List<string>();
            for (var col = 0; col < columns; col++)
            {
                list.Add((table.Cells[0, col].TextString ?? string.Empty).Trim());
            }

            return list;
        }

        private static string ResolveCrossingKey(Cell cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            var text = (cell.TextString ?? string.Empty).Trim();
            if (!string.IsNullOrEmpty(text))
            {
                return text;
            }

            try
            {
                text = (cell.BlockAttributeValue ?? string.Empty).Trim();
            }
            catch
            {
                text = string.Empty;
            }

            return text;
        }

        private static void SetCellValue(Cell cell, string value)
        {
            if (cell == null)
            {
                return;
            }

            cell.TextString = value ?? string.Empty;
            try
            {
                cell.BlockAttributeValue = value ?? string.Empty;
            }
            catch
            {
                // Some cells may not be block-based; ignore failures.
            }
        }

        private void UpdateMainTable(Table table, IDictionary<string, CrossingRecord> byKey)
        {
            for (var row = 1; row < table.Rows.Count; row++)
            {
                var cell = table.Cells[row, 0];
                var crossingKey = ResolveCrossingKey(cell);
                if (string.IsNullOrEmpty(crossingKey))
                {
                    continue;
                }

                var canonicalKey = crossingKey.Trim().ToUpperInvariant();

                CrossingRecord record;
                if (!byKey.TryGetValue(canonicalKey, out record))
                {
                    // Try numeric-only match for flexibility.
                    record = byKey.Values.FirstOrDefault(r => CrossingRecord.CompareCrossingKeys(r.Crossing, crossingKey) == 0);
                }

                if (record == null)
                {
                    continue;
                }

                SetCellValue(table.Cells[row, 0], record.Crossing);
                SetCellValue(table.Cells[row, 1], record.Owner);
                SetCellValue(table.Cells[row, 2], record.Description);
                SetCellValue(table.Cells[row, 3], record.Location);
                SetCellValue(table.Cells[row, 4], record.DwgRef);
            }
        }

        private void UpdatePageTable(Table table, IDictionary<string, CrossingRecord> byKey)
        {
            for (var row = 1; row < table.Rows.Count; row++)
            {
                var cell = table.Cells[row, 0];
                var crossingKey = ResolveCrossingKey(cell);
                if (string.IsNullOrEmpty(crossingKey))
                {
                    continue;
                }

                CrossingRecord record;
                if (!byKey.TryGetValue(crossingKey.Trim().ToUpperInvariant(), out record))
                {
                    record = byKey.Values.FirstOrDefault(r => CrossingRecord.CompareCrossingKeys(r.Crossing, crossingKey) == 0);
                }

                if (record == null)
                {
                    continue;
                }

                SetCellValue(table.Cells[row, 0], record.Crossing);
                SetCellValue(table.Cells[row, 1], record.Owner);
                SetCellValue(table.Cells[row, 2], record.Description);
            }
        }

        private void UpdateLatLongTable(Table table, IDictionary<string, CrossingRecord> byKey)
        {
            // Only expect a single data row but iterate defensively.
            for (var row = 1; row < table.Rows.Count; row++)
            {
                var cell = table.Cells[row, 0];
                var crossingKey = ResolveCrossingKey(cell);
                if (string.IsNullOrEmpty(crossingKey))
                {
                    continue;
                }

                CrossingRecord record;
                if (!byKey.TryGetValue(crossingKey.Trim().ToUpperInvariant(), out record))
                {
                    record = byKey.Values.FirstOrDefault(r => CrossingRecord.CompareCrossingKeys(r.Crossing, crossingKey) == 0);
                }

                if (record == null)
                {
                    continue;
                }

                SetCellValue(table.Cells[row, 0], record.Crossing);
                SetCellValue(table.Cells[row, 1], record.Description);
                SetCellValue(table.Cells[row, 2], record.Lat);
                SetCellValue(table.Cells[row, 3], record.Long);
            }
        }
    }
}
