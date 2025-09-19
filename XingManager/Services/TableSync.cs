using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
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

                foreach (ObjectId btrId in blockTable)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    foreach (ObjectId entId in btr)
                    {
                        var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        var table = ent as Table;
                        if (table == null)
                        {
                            continue;
                        }

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
                var expectedValue = NormalizeHeader(expected[i], i);
                if (!string.Equals(headers[i], expectedValue, StringComparison.Ordinal))
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
                list.Add(NormalizeHeader(table.Cells[0, col].TextString, col));
            }

            return list;
        }

        private static string NormalizeHeader(string header, int columnIndex)
        {
            if (header == null)
            {
                return string.Empty;
            }

            var builder = new StringBuilder(header.Length);
            foreach (var ch in header)
            {
                if (ch == ' ' || ch == '.' || ch == ',' || ch == '#' || ch == '_')
                {
                    continue;
                }

                builder.Append(char.ToUpperInvariant(ch));
            }

            var normalized = builder.ToString();
            if (columnIndex == 4)
            {
                if (normalized == "DWGREF" || normalized == "XINGDWGREF" || normalized == "XINGDWGREFNO" || normalized == "XINGDWGREFNUMBER")
                {
                    return "DWGREF";
                }
            }

            return normalized;
        }

        private static string ResolveCrossingKey(Table table, int row, int col)
        {
            if (table == null)
            {
                return string.Empty;
            }

            Cell cell;
            try
            {
                cell = table.Cells[row, col];
            }
            catch
            {
                return string.Empty;
            }

            var text = cell?.TextString;
            if (!string.IsNullOrWhiteSpace(text))
            {
                return text.Trim();
            }

            var blockValue = TryGetBlockAttributeValue(table, row, col, "CROSSING");
            if (!string.IsNullOrWhiteSpace(blockValue))
            {
                return blockValue.Trim();
            }

            return string.Empty;
        }

        private static void SetCellCrossingValue(Table t, int row, int col, string crossingText)
        {
            if (!TrySetBlockAttributeValue(t, row, col, "CROSSING", crossingText))
            {
                var cell = t.Cells[row, col];
                cell.TextString = crossingText ?? string.Empty;
            }
        }

        private static void SetCellValue(Cell cell, string value)
        {
            if (cell == null)
            {
                return;
            }

            cell.TextString = value ?? string.Empty;
        }

        private void UpdateMainTable(Table table, IDictionary<string, CrossingRecord> byKey)
        {
            for (var row = 1; row < table.Rows.Count; row++)
            {
                var crossingKey = ResolveCrossingKey(table, row, 0);
                var record = FindRecordForKey(byKey, crossingKey);
                if (record == null)
                {
                    continue;
                }

                SetCellCrossingValue(table, row, 0, record.Crossing);
                SetCellValue(table.Cells[row, 1], record.Owner);
                SetCellValue(table.Cells[row, 2], record.Description);
                SetCellValue(table.Cells[row, 3], record.Location);
                SetCellValue(table.Cells[row, 4], record.DwgRef);
            }

            RefreshTable(table);
        }

        private void UpdatePageTable(Table table, IDictionary<string, CrossingRecord> byKey)
        {
            for (var row = 1; row < table.Rows.Count; row++)
            {
                var crossingKey = ResolveCrossingKey(table, row, 0);
                var record = FindRecordForKey(byKey, crossingKey);
                if (record == null)
                {
                    continue;
                }

                SetCellCrossingValue(table, row, 0, record.Crossing);
                SetCellValue(table.Cells[row, 1], record.Owner);
                SetCellValue(table.Cells[row, 2], record.Description);
            }

            RefreshTable(table);
        }

        private void UpdateLatLongTable(Table table, IDictionary<string, CrossingRecord> byKey)
        {
            // Only expect a single data row but iterate defensively.
            for (var row = 1; row < table.Rows.Count; row++)
            {
                var crossingKey = ResolveCrossingKey(table, row, 0);
                var record = FindRecordForKey(byKey, crossingKey);
                if (record == null)
                {
                    continue;
                }

                SetCellCrossingValue(table, row, 0, record.Crossing);
                SetCellValue(table.Cells[row, 1], record.Description);
                SetCellValue(table.Cells[row, 2], record.Lat);
                SetCellValue(table.Cells[row, 3], record.Long);
            }

            RefreshTable(table);
        }

        private static CrossingRecord FindRecordForKey(IDictionary<string, CrossingRecord> byKey, string crossingKey)
        {
            if (byKey == null || string.IsNullOrWhiteSpace(crossingKey))
            {
                return null;
            }

            var trimmedKey = crossingKey.Trim();

            CrossingRecord record;
            if (byKey.TryGetValue(trimmedKey, out record))
            {
                return record;
            }

            record = byKey.Values.FirstOrDefault(r => CrossingRecord.CompareCrossingKeys(r.Crossing, trimmedKey) == 0);
            return record;
        }

        private static string TryGetBlockAttributeValue(Table table, int row, int col, string tag)
        {
            if (table == null || string.IsNullOrEmpty(tag))
            {
                return string.Empty;
            }

            const string methodName = "GetBlockAttributeValue";
            var type = table.GetType();
            var methods = type.GetMethods().Where(m => string.Equals(m.Name, methodName, StringComparison.Ordinal));

            foreach (var method in methods)
            {
                var parameters = method.GetParameters();
                if (parameters.Length < 3)
                {
                    continue;
                }

                if (!parameters[2].ParameterType.IsAssignableFrom(typeof(string)))
                {
                    continue;
                }

                var args = new object[parameters.Length];
                if (!TryConvertParameter(row, parameters[0], out args[0]) ||
                    !TryConvertParameter(col, parameters[1], out args[1]) ||
                    !TryConvertParameter(tag, parameters[2], out args[2]))
                {
                    continue;
                }

                var skip = false;
                for (var i = 3; i < parameters.Length; i++)
                {
                    var parameter = parameters[i];
                    if (!parameter.IsOptional)
                    {
                        skip = true;
                        break;
                    }
                    args[i] = Type.Missing;
                }

                if (skip)
                {
                    continue;
                }

                try
                {
                    var result = method.Invoke(table, args);
                    if (result == null)
                    {
                        continue;
                    }

                    var text = Convert.ToString(result, CultureInfo.InvariantCulture);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        return text;
                    }
                }
                catch
                {
                }
            }

            return string.Empty;
        }

        private static bool TrySetBlockAttributeValue(Table table, int row, int col, string tag, string value)
        {
            if (table == null || string.IsNullOrEmpty(tag))
            {
                return false;
            }

            const string methodName = "SetBlockAttributeValue";
            var type = table.GetType();
            var methods = type.GetMethods().Where(m => string.Equals(m.Name, methodName, StringComparison.Ordinal));

            foreach (var method in methods)
            {
                var parameters = method.GetParameters();
                if (parameters.Length < 4)
                {
                    continue;
                }

                if (!parameters[2].ParameterType.IsAssignableFrom(typeof(string)) ||
                    !parameters[3].ParameterType.IsAssignableFrom(typeof(string)))
                {
                    continue;
                }

                var args = new object[parameters.Length];
                if (!TryConvertParameter(row, parameters[0], out args[0]) ||
                    !TryConvertParameter(col, parameters[1], out args[1]) ||
                    !TryConvertParameter(tag, parameters[2], out args[2]) ||
                    !TryConvertParameter(value ?? string.Empty, parameters[3], out args[3]))
                {
                    continue;
                }

                var skip = false;
                for (var i = 4; i < parameters.Length; i++)
                {
                    var parameter = parameters[i];
                    if (!parameter.IsOptional)
                    {
                        skip = true;
                        break;
                    }
                    args[i] = Type.Missing;
                }

                if (skip)
                {
                    continue;
                }

                try
                {
                    method.Invoke(table, args);
                    return true;
                }
                catch
                {
                }
            }

            return false;
        }

        private static bool TryConvertParameter(object value, ParameterInfo parameter, out object converted)
        {
            var targetType = parameter.ParameterType;
            var underlying = Nullable.GetUnderlyingType(targetType) ?? targetType;

            try
            {
                if (value == null)
                {
                    if (!underlying.IsValueType || underlying == typeof(string))
                    {
                        converted = null;
                        return true;
                    }

                    converted = Activator.CreateInstance(underlying);
                    return true;
                }

                if (underlying.IsInstanceOfType(value))
                {
                    converted = value;
                    return true;
                }

                converted = Convert.ChangeType(value, underlying, CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
                converted = null;
                return false;
            }
        }

        private static void RefreshTable(Table table)
        {
            if (table == null)
            {
                return;
            }

            var recompute = table.GetType().GetMethod("RecomputeTableBlock", new[] { typeof(bool) });
            if (recompute != null)
            {
                try
                {
                    recompute.Invoke(table, new object[] { true });
                    return;
                }
                catch
                {
                }
            }

            try
            {
                table.GenerateLayout();
            }
            catch
            {
            }
        }
    }
}
