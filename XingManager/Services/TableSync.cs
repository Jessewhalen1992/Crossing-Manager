using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
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

        private static readonly HashSet<string> MainDwgColumnSynonyms = new HashSet<string>(StringComparer.Ordinal)
        {
            "DWGREF",
            "DWGREFNO",
            "DWGREFNUMBER",
            "XINGDWGREF",
            "XINGDWGREFNO",
            "XINGDWGREFNUMBER"
        };

        private static readonly string[] CrossingAttributeTags =
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
        };

        private static readonly Regex MTextFormattingCommandRegex = new Regex("\\\\[^;\\\\{}]*;", RegexOptions.Compiled);
        private static readonly Regex MTextResidualCommandRegex = new Regex("\\\\[^{}]", RegexOptions.Compiled);
        private static readonly Regex MTextSpecialCodeRegex = new Regex("%%[^\\s]+", RegexOptions.Compiled);

        private readonly TableFactory _factory;
        private Editor _ed;

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
            _ed = doc.Editor;
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
                        var typeLabel = type.ToString().ToUpperInvariant();

                        if (type == XingTableType.Unknown)
                        {
                            Log(string.Format(CultureInfo.InvariantCulture, "Table {0}: {1} matched=0 updated=0", entId.Handle, typeLabel));
                            var headerLog = BuildHeaderLog(table);
                            if (!string.IsNullOrEmpty(headerLog))
                            {
                                Log(string.Format(CultureInfo.InvariantCulture, "Table {0}: {1} {2}", entId.Handle, typeLabel, headerLog));
                            }

                            continue;
                        }

                        table.UpgradeOpen();
                        var matched = 0;
                        var updated = 0;

                        try
                        {
                            switch (type)
                            {
                                case XingTableType.Main:
                                    UpdateMainTable(table, byKey, out matched, out updated);
                                    break;
                                case XingTableType.Page:
                                    UpdatePageTable(table, byKey, out matched, out updated);
                                    break;
                                case XingTableType.LatLong:
                                    UpdateLatLongTable(table, byKey, out matched, out updated);
                                    break;
                            }

                            _factory.TagTable(tr, table, typeLabel);
                            Log(string.Format(CultureInfo.InvariantCulture, "Table {0}: {1} matched={2} updated={3}", entId.Handle, typeLabel, matched, updated));
                        }
                        catch (System.Exception ex)
                        {
                            Log(string.Format(CultureInfo.InvariantCulture, "Failed to update table {0}: {1}", entId.Handle, ex.Message));
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

        private void Log(string msg)
        {
            try
            {
                _ed?.WriteMessage("\n[CrossingManager] " + msg);
            }
            catch
            {
            }
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

            header = StripMTextFormatting(header);

            var builder = new StringBuilder(header.Length);
            foreach (var ch in header)
            {
                if (char.IsWhiteSpace(ch) || ch == '.' || ch == ',' || ch == '#' || ch == '_' || ch == '-' || ch == '/' || ch == '(' || ch == ')' || ch == '%' || ch == '|')
                {
                    continue;
                }

                builder.Append(char.ToUpperInvariant(ch));
            }

            var normalized = builder.ToString();
            if (columnIndex == 4)
            {
                if (MainDwgColumnSynonyms.Contains(normalized))
                {
                    return "DWGREF";
                }
            }

            return normalized;
        }

        private static string StripMTextFormatting(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var withoutCommands = MTextFormattingCommandRegex.Replace(value, string.Empty);
            var withoutResidual = MTextResidualCommandRegex.Replace(withoutCommands, string.Empty);
            var withoutSpecial = MTextSpecialCodeRegex.Replace(withoutResidual, string.Empty);
            return withoutSpecial.Replace("{", string.Empty).Replace("}", string.Empty);
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

            foreach (var tag in CrossingAttributeTags)
            {
                var blockValue = TryGetBlockAttributeValue(table, row, col, tag);
                if (!string.IsNullOrWhiteSpace(blockValue))
                {
                    return blockValue.Trim();
                }
            }

            var attrProperty = cell?.GetType().GetProperty("BlockAttributeValue", BindingFlags.Public | BindingFlags.Instance);
            if (attrProperty != null)
            {
                try
                {
                    var attrValue = attrProperty.GetValue(cell, null) as string;
                    if (!string.IsNullOrWhiteSpace(attrValue))
                    {
                        return attrValue.Trim();
                    }
                }
                catch
                {
                }
            }

            foreach (var content in EnumerateCellContents(cell))
            {
                if (!string.IsNullOrWhiteSpace(content))
                {
                    return content.Trim();
                }
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

        private void UpdateMainTable(Table table, IDictionary<string, CrossingRecord> byKey, out int matched, out int updated)
        {
            matched = 0;
            updated = 0;
            for (var row = 1; row < table.Rows.Count; row++)
            {
                var crossingKey = ResolveCrossingKey(table, row, 0);
                var record = FindRecordForKey(byKey, crossingKey);
                if (record == null)
                {
                    continue;
                }

                matched++;

                var desiredCrossing = (record.Crossing ?? string.Empty).Trim();
                var desiredOwner = record.Owner ?? string.Empty;
                var desiredDescription = record.Description ?? string.Empty;
                var desiredLocation = record.Location ?? string.Empty;
                var desiredDwgRef = record.DwgRef ?? string.Empty;

                var rowUpdated = false;

                if (ValueDiffers(crossingKey, desiredCrossing))
                {
                    rowUpdated = true;
                }

                var ownerCell = table.Cells[row, 1];
                if (ValueDiffers(ReadCellText(ownerCell), desiredOwner))
                {
                    rowUpdated = true;
                }

                var descriptionCell = table.Cells[row, 2];
                if (ValueDiffers(ReadCellText(descriptionCell), desiredDescription))
                {
                    rowUpdated = true;
                }

                var locationCell = table.Cells[row, 3];
                if (ValueDiffers(ReadCellText(locationCell), desiredLocation))
                {
                    rowUpdated = true;
                }

                var dwgRefCell = table.Cells[row, 4];
                if (ValueDiffers(ReadCellText(dwgRefCell), desiredDwgRef))
                {
                    rowUpdated = true;
                }

                SetCellCrossingValue(table, row, 0, desiredCrossing);
                SetCellValue(ownerCell, desiredOwner);
                SetCellValue(descriptionCell, desiredDescription);
                SetCellValue(locationCell, desiredLocation);
                SetCellValue(dwgRefCell, desiredDwgRef);

                if (rowUpdated)
                {
                    updated++;
                }
            }

            RefreshTable(table);
        }

        private void UpdatePageTable(Table table, IDictionary<string, CrossingRecord> byKey, out int matched, out int updated)
        {
            matched = 0;
            updated = 0;
            for (var row = 1; row < table.Rows.Count; row++)
            {
                var crossingKey = ResolveCrossingKey(table, row, 0);
                var record = FindRecordForKey(byKey, crossingKey);
                if (record == null)
                {
                    continue;
                }

                matched++;

                var desiredCrossing = (record.Crossing ?? string.Empty).Trim();
                var desiredOwner = record.Owner ?? string.Empty;
                var desiredDescription = record.Description ?? string.Empty;

                var rowUpdated = false;

                if (ValueDiffers(crossingKey, desiredCrossing))
                {
                    rowUpdated = true;
                }

                var ownerCell = table.Cells[row, 1];
                if (ValueDiffers(ReadCellText(ownerCell), desiredOwner))
                {
                    rowUpdated = true;
                }

                var descriptionCell = table.Cells[row, 2];
                if (ValueDiffers(ReadCellText(descriptionCell), desiredDescription))
                {
                    rowUpdated = true;
                }

                SetCellCrossingValue(table, row, 0, desiredCrossing);
                SetCellValue(ownerCell, desiredOwner);
                SetCellValue(descriptionCell, desiredDescription);

                if (rowUpdated)
                {
                    updated++;
                }
            }

            RefreshTable(table);
        }

        private void UpdateLatLongTable(Table table, IDictionary<string, CrossingRecord> byKey, out int matched, out int updated)
        {
            matched = 0;
            updated = 0;
            // Only expect a single data row but iterate defensively.
            for (var row = 1; row < table.Rows.Count; row++)
            {
                var crossingKey = ResolveCrossingKey(table, row, 0);
                var record = FindRecordForKey(byKey, crossingKey);
                if (record == null)
                {
                    continue;
                }

                matched++;

                var desiredCrossing = (record.Crossing ?? string.Empty).Trim();
                var desiredDescription = record.Description ?? string.Empty;
                var desiredLat = record.Lat ?? string.Empty;
                var desiredLong = record.Long ?? string.Empty;

                var rowUpdated = false;

                if (ValueDiffers(crossingKey, desiredCrossing))
                {
                    rowUpdated = true;
                }

                var descriptionCell = table.Cells[row, 1];
                if (ValueDiffers(ReadCellText(descriptionCell), desiredDescription))
                {
                    rowUpdated = true;
                }

                var latCell = table.Cells[row, 2];
                if (ValueDiffers(ReadCellText(latCell), desiredLat))
                {
                    rowUpdated = true;
                }

                var longCell = table.Cells[row, 3];
                if (ValueDiffers(ReadCellText(longCell), desiredLong))
                {
                    rowUpdated = true;
                }

                SetCellCrossingValue(table, row, 0, desiredCrossing);
                SetCellValue(descriptionCell, desiredDescription);
                SetCellValue(latCell, desiredLat);
                SetCellValue(longCell, desiredLong);

                if (rowUpdated)
                {
                    updated++;
                }
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

            var normalizedKey = NormalizeCrossingLookupKey(trimmedKey);
            if (!string.IsNullOrEmpty(normalizedKey) && !string.Equals(normalizedKey, trimmedKey, StringComparison.Ordinal))
            {
                if (byKey.TryGetValue(normalizedKey, out record))
                {
                    return record;
                }
            }

            record = byKey.Values.FirstOrDefault(r => string.Equals(NormalizeCrossingLookupKey(r.Crossing), normalizedKey, StringComparison.Ordinal) || CrossingRecord.CompareCrossingKeys(r.Crossing, trimmedKey) == 0);
            return record;
        }

        private static string NormalizeCrossingLookupKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return string.Empty;
            }

            var trimmed = key.Trim().ToUpperInvariant();
            var builder = new StringBuilder(trimmed.Length);
            var inDigits = false;
            var appendedDigitInRun = false;
            var hasAnyDigits = trimmed.Any(char.IsDigit);

            foreach (var ch in trimmed)
            {
                if (char.IsDigit(ch))
                {
                    if (!inDigits)
                    {
                        inDigits = true;
                        appendedDigitInRun = false;
                    }

                    if (ch == '0' && !appendedDigitInRun)
                    {
                        continue;
                    }

                    builder.Append(ch);
                    appendedDigitInRun = true;
                    continue;
                }

                if (inDigits)
                {
                    if (!appendedDigitInRun && hasAnyDigits)
                    {
                        builder.Append('0');
                    }

                    inDigits = false;
                    appendedDigitInRun = false;
                }

                if (char.IsLetter(ch))
                {
                    builder.Append(ch);
                }
            }

            if (inDigits && !appendedDigitInRun)
            {
                builder.Append('0');
            }

            return builder.ToString();
        }

        private static bool ValueDiffers(string existing, string desired)
        {
            var left = (existing ?? string.Empty).Trim();
            var right = (desired ?? string.Empty).Trim();
            return !string.Equals(left, right, StringComparison.Ordinal);
        }

        private static string ReadCellText(Cell cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            try
            {
                return cell.TextString ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static IEnumerable<string> EnumerateCellContents(Cell cell)
        {
            if (cell == null)
            {
                yield break;
            }

            var contentsProperty = cell.GetType().GetProperty("Contents", BindingFlags.Public | BindingFlags.Instance);
            if (contentsProperty == null)
            {
                yield break;
            }

            object contents;
            try
            {
                contents = contentsProperty.GetValue(cell, null);
            }
            catch
            {
                yield break;
            }

            var enumerable = contents as IEnumerable;
            if (enumerable == null)
            {
                yield break;
            }

            foreach (var item in enumerable)
            {
                if (item == null)
                {
                    continue;
                }

                var textProp = item.GetType().GetProperty("TextString", BindingFlags.Public | BindingFlags.Instance);
                if (textProp != null)
                {
                    string text = null;
                    try
                    {
                        text = textProp.GetValue(item, null) as string;
                    }
                    catch
                    {
                        text = null;
                    }

                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        yield return text;
                        continue;
                    }
                }

                var textValue = item.ToString();
                if (!string.IsNullOrWhiteSpace(textValue))
                {
                    yield return textValue;
                }
            }
        }

        private static string BuildHeaderLog(Table table)
        {
            if (table == null)
            {
                return "headers=[]";
            }

            var rowCount = table.Rows.Count;
            var columnCount = table.Columns.Count;

            if (rowCount == 0 || columnCount <= 0)
            {
                return "headers=[]";
            }

            var headers = new List<string>(columnCount);
            for (var col = 0; col < columnCount; col++)
            {
                string text;
                try
                {
                    text = table.Cells[0, col].TextString ?? string.Empty;
                }
                catch
                {
                    text = string.Empty;
                }

                headers.Add(text.Trim());
            }

            return "headers=[" + string.Join("|", headers) + "]";
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
