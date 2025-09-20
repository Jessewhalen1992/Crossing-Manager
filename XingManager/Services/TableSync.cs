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

        private const int MaxHeaderRowsToScan = 5;

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
            if (doc == null) throw new ArgumentNullException("doc");
            if (records == null) throw new ArgumentNullException("records");

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
                        if (table == null) continue;

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
            if (table == null) return XingTableType.Unknown;

            // Prefer explicit tag if present
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
                                if (string.Equals(text, "MAIN", StringComparison.OrdinalIgnoreCase)) return XingTableType.Main;
                                if (string.Equals(text, "PAGE", StringComparison.OrdinalIgnoreCase)) return XingTableType.Page;
                                if (string.Equals(text, "LATLONG", StringComparison.OrdinalIgnoreCase)) return XingTableType.LatLong;
                            }
                        }
                    }
                }
            }

            // Header-free heuristics: MAIN=5 cols, PAGE=3 cols, LATLONG=4 cols (+ sanity)
            if (table.Columns.Count == 5) return XingTableType.Main;
            if (table.Columns.Count == 3) return XingTableType.Page;

            if (table.Columns.Count == 4 && table.Rows.Count >= 1)
            {
                if (HasHeaderRow(table, 4, IsLatLongHeader) || LooksLikeLatLongTable(table))
                    return XingTableType.LatLong;
            }

            return XingTableType.Unknown;
        }

        private static int CountXKeys(Table t, int keyCol)
        {
            var rows = t?.Rows.Count ?? 0;
            var count = 0;
            for (var r = 0; r < rows; r++)
            {
                var key = NormalizeKeyForLookup(ResolveCrossingKey(t, r, keyCol));
                if (!string.IsNullOrEmpty(key) && key[0] == 'X' && key.Length > 1 && key.Skip(1).All(char.IsDigit))
                {
                    count++;
                }
            }

            return count;
        }

        private void Log(string msg)
        {
            try { _ed?.WriteMessage("\n[CrossingManager] " + msg); } catch { }
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
            if (headers == null || headers.Count != expected.Length) return false;
            for (var i = 0; i < expected.Length; i++)
            {
                var expectedValue = NormalizeHeader(expected[i], i);
                if (!string.Equals(headers[i], expectedValue, StringComparison.Ordinal))
                    return false;
            }
            return true;
        }

        private static bool HasHeaderRow(Table table, int columnCount, Func<List<string>, bool> predicate)
        {
            int headerRowIndex;
            return TryFindHeaderRow(table, columnCount, predicate, out headerRowIndex);
        }

        private static int GetDataStartRow(Table table, int columnCount, Func<List<string>, bool> predicate)
        {
            int headerRowIndex;
            if (TryFindHeaderRow(table, columnCount, predicate, out headerRowIndex))
            {
                return Math.Min(headerRowIndex + 1, table?.Rows.Count ?? 0);
            }
            return 0;
        }

        private static bool TryFindHeaderRow(Table table, int columnCount, Func<List<string>, bool> predicate, out int headerRowIndex)
        {
            headerRowIndex = -1;
            if (table == null || predicate == null || columnCount <= 0) return false;
            var rowCount = table.Rows.Count;
            if (rowCount <= 0) return false;

            var rowsToScan = Math.Min(rowCount, MaxHeaderRowsToScan);
            for (var row = 0; row < rowsToScan; row++)
            {
                var headers = ReadHeaders(table, columnCount, row);
                if (predicate(headers))
                {
                    headerRowIndex = row;
                    return true;
                }
            }
            return false;
        }

        private static bool LooksLikeLatLongTable(Table table)
        {
            if (table == null) return false;

            var rowCount = table.Rows.Count;
            if (rowCount <= 0 || table.Columns.Count != 4) return false;

            var rowsToScan = Math.Min(rowCount, MaxHeaderRowsToScan);
            var candidates = 0;

            for (var row = 0; row < rowsToScan; row++)
            {
                var latText = ReadCellText(table, row, 2);
                var longText = ReadCellText(table, row, 3);

                if (IsCoordinateValue(latText, -90.0, 90.0) && IsCoordinateValue(longText, -180.0, 180.0))
                {
                    var crossing = ReadCellText(table, row, 0);
                    if (string.IsNullOrWhiteSpace(crossing)) return false;
                    candidates++;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(latText) && string.IsNullOrWhiteSpace(longText))
                {
                    continue;
                }

                var normalizedLat = NormalizeHeader(latText, 2);
                var normalizedLong = NormalizeHeader(longText, 3);
                if (normalizedLat.StartsWith("LAT", StringComparison.Ordinal) && normalizedLong.StartsWith("LONG", StringComparison.Ordinal))
                {
                    continue;
                }

                return false;
            }

            return candidates > 0;
        }

        private static List<string> ReadHeaders(Table table, int columns, int rowIndex)
        {
            var list = new List<string>(columns);
            for (var col = 0; col < columns; col++)
            {
                string text = string.Empty;
                if (table != null && rowIndex >= 0 && rowIndex < table.Rows.Count && col < table.Columns.Count)
                {
                    try { text = table.Cells[rowIndex, col].TextString ?? string.Empty; }
                    catch { text = string.Empty; }
                }
                list.Add(NormalizeHeader(text, col));
            }
            return list;
        }

        private static string NormalizeHeader(string header, int columnIndex)
        {
            if (header == null) return string.Empty;

            header = StripMTextFormatting(header);

            var builder = new StringBuilder(header.Length);
            foreach (var ch in header)
            {
                if (char.IsWhiteSpace(ch) || ch == '.' || ch == ',' || ch == '#' || ch == '_' || ch == '-' || ch == '/' || ch == '(' || ch == ')' || ch == '%' || ch == '|')
                    continue;
                builder.Append(char.ToUpperInvariant(ch));
            }

            var normalized = builder.ToString();
            if (columnIndex == 4)
            {
                if (MainDwgColumnSynonyms.Contains(normalized)) return "DWGREF";
            }
            return normalized;
        }

        private static string StripMTextFormatting(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            var withoutCommands = MTextFormattingCommandRegex.Replace(value, string.Empty);
            var withoutResidual = MTextResidualCommandRegex.Replace(withoutCommands, string.Empty);
            var withoutSpecial = MTextSpecialCodeRegex.Replace(withoutResidual, string.Empty);
            return withoutSpecial.Replace("{", string.Empty).Replace("}", string.Empty);
        }

        internal static string ResolveCrossingKey(Table table, int row, int col)
        {
            if (table == null) return string.Empty;

            Cell cell;
            try { cell = table.Cells[row, col]; }
            catch { cell = null; }

            string directText = null;
            try { directText = cell?.TextString; } catch { directText = null; }

            var cleanedDirect = CleanCellText(directText);
            if (!string.IsNullOrWhiteSpace(cleanedDirect)) return cleanedDirect;

            foreach (var tag in CrossingAttributeTags)
            {
                var blockVal = TableCellProbe.TryGetCellBlockAttr(table, row, col, tag);
                var cleaned = CleanCellText(blockVal);
                if (!string.IsNullOrWhiteSpace(cleaned)) return cleaned;
            }

            foreach (var tag in CrossingAttributeTags)
            {
                var blockValue = TryGetBlockAttributeValue(table, row, col, tag);
                var cleanedBlockValue = CleanCellText(blockValue);
                if (!string.IsNullOrWhiteSpace(cleanedBlockValue)) return cleanedBlockValue;
            }

            var attrProperty = cell?.GetType().GetProperty("BlockAttributeValue", BindingFlags.Public | BindingFlags.Instance);
            if (attrProperty != null)
            {
                try
                {
                    var attrValue = attrProperty.GetValue(cell, null) as string;
                    var cleanedAttrValue = CleanCellText(attrValue);
                    if (!string.IsNullOrWhiteSpace(cleanedAttrValue)) return cleanedAttrValue;
                }
                catch { }
            }

            foreach (var content in EnumerateCellContents(cell))
            {
                var cleanedContent = CleanCellText(content);
                if (!string.IsNullOrWhiteSpace(cleanedContent)) return cleanedContent;
            }

            // LAST RESORT: discover & try any attribute tags present on the block content in this cell
            foreach (var discoveredTag in GetBlockAttributeTagsFromCell(table, row, col))
            {
                var any = TryGetBlockAttributeValue(table, row, col, discoveredTag);
                var cleaned = CleanCellText(any);
                if (!string.IsNullOrWhiteSpace(cleaned)) return cleaned;
            }

            return string.Empty;
        }

        private static string CleanCellText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            return StripMTextFormatting(text).Trim();
        }

        private static string Norm(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            return StripMTextFormatting(s).Trim();
        }

        private static string ReadNorm(Table t, int row, int col)
        {
            return Norm(ReadCellText(t, row, col));
        }

        private static CrossingRecord FindRecordByMainColumns(Table t, int row, IEnumerable<CrossingRecord> records)
        {
            var owner = ReadNorm(t, row, 1);
            var desc = ReadNorm(t, row, 2);
            var loc = ReadNorm(t, row, 3);
            var dwg = ReadNorm(t, row, 4);

            var candidates = records.Where(r =>
                string.Equals(Norm(r.Owner), owner, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Description), desc, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Location), loc, StringComparison.Ordinal) &&
                string.Equals(Norm(r.DwgRef), dwg, StringComparison.Ordinal)).ToList();
            if (candidates.Count == 1) return candidates[0];

            candidates = records.Where(r =>
                string.Equals(Norm(r.Description), desc, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Location), loc, StringComparison.Ordinal) &&
                string.Equals(Norm(r.DwgRef), dwg, StringComparison.Ordinal)).ToList();
            if (candidates.Count == 1) return candidates[0];

            candidates = records.Where(r =>
                string.Equals(Norm(r.Description), desc, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Location), loc, StringComparison.Ordinal)).ToList();
            if (candidates.Count == 1) return candidates[0];

            candidates = records.Where(r => string.Equals(Norm(r.Description), desc, StringComparison.Ordinal)).ToList();
            return candidates.Count == 1 ? candidates[0] : null;
        }

        private static CrossingRecord FindRecordByPageColumns(Table t, int row, IEnumerable<CrossingRecord> records)
        {
            var owner = ReadNorm(t, row, 1);
            var desc = ReadNorm(t, row, 2);

            var candidates = records.Where(r =>
                string.Equals(Norm(r.Owner), owner, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Description), desc, StringComparison.Ordinal)).ToList();
            if (candidates.Count == 1) return candidates[0];

            candidates = records.Where(r => string.Equals(Norm(r.Description), desc, StringComparison.Ordinal)).ToList();
            return candidates.Count == 1 ? candidates[0] : null;
        }

        private static CrossingRecord FindRecordByLatLongColumns(Table t, int row, IEnumerable<CrossingRecord> records)
        {
            var desc = ReadNorm(t, row, 1);
            var lat = ReadNorm(t, row, 2);
            var lng = ReadNorm(t, row, 3);

            var candidates = records.Where(r =>
                string.Equals(Norm(r.Description), desc, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Lat), lat, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Long), lng, StringComparison.Ordinal)).ToList();
            if (candidates.Count == 1) return candidates[0];

            candidates = records.Where(r => string.Equals(Norm(r.Description), desc, StringComparison.Ordinal)).ToList();
            return candidates.Count == 1 ? candidates[0] : null;
        }

        internal static string NormalizeKeyForLookup(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;

            s = StripMTextFormatting(s).Trim().ToUpperInvariant();
            var sb = new StringBuilder(s.Length);
            foreach (var ch in s)
            {
                if (char.IsLetterOrDigit(ch)) sb.Append(ch);
            }

            var collapsed = sb.ToString();
            var match = System.Text.RegularExpressions.Regex.Match(collapsed, @"^X0*(\d+)$");
            if (!match.Success) return string.Empty;

            var digits = match.Groups[1].Value.TrimStart('0');
            if (digits.Length == 0) digits = "0";

            return "X" + digits;
        }

        private static void SetCellCrossingValue(Table t, int row, int col, string crossingText)
        {
            if (t == null) return;

            // 1) Try to set the block attribute by any of our known tags
            //    (this covers LABEL, X_NO, NUMBER, etc., not just CROSSING)
            foreach (var tag in CrossingAttributeTags)
            {
                if (TrySetBlockAttributeValue(t, row, col, tag, crossingText))
                    return;
            }

            // 2) If the cell currently hosts a block, do NOT overwrite it with text.
            Cell cell = null;
            try { cell = t.Cells[row, col]; } catch { cell = null; }

            if (CellHasBlockContent(cell))
            {
                // We couldn't set an attribute; leave the block intact.
                // (Optional: add a log if you want to see which rows failed.)
                return;
            }

            // 3) Only fall back to plain text when the cell is not a block cell
            try
            {
                if (cell != null) cell.TextString = crossingText ?? string.Empty;
            }
            catch { }
        }

        private static void SetCellValue(Cell cell, string value)
        {
            if (cell == null) return;
            cell.TextString = value ?? string.Empty;
        }

        private void UpdateMainTable(Table table, IDictionary<string, CrossingRecord> byKey, out int matched, out int updated)
        {
            matched = 0;
            updated = 0;
            if (table == null) return;

            var columnCount = table.Columns.Count;
            var records = byKey?.Values ?? Enumerable.Empty<CrossingRecord>();

            for (var row = 0; row < table.Rows.Count; row++)
            {
                var rawKey = ResolveCrossingKey(table, row, 0);
                var key = NormalizeKeyForLookup(rawKey);
                CrossingRecord record = null;

                if (!string.IsNullOrEmpty(key) && byKey != null)
                {
                    if (!byKey.TryGetValue(key, out record))
                    {
                        record = byKey.Values.FirstOrDefault(r => CrossingRecord.CompareCrossingKeys(r.Crossing, key) == 0);
                    }
                }

                if (record == null)
                {
                    record = FindRecordByMainColumns(table, row, records);
                }

                if (record == null)
                {
                    Log(string.Format(CultureInfo.InvariantCulture, "Row {0} -> NO MATCH (key='{1}')", row, rawKey));
                    continue;
                }

                var logKey = !string.IsNullOrEmpty(key) ? key : (rawKey ?? string.Empty);
                matched++;

                var desiredCrossing = (record.Crossing ?? string.Empty).Trim();
                var desiredOwner = record.Owner ?? string.Empty;
                var desiredDescription = record.Description ?? string.Empty;
                var desiredLocation = record.Location ?? string.Empty;
                var desiredDwgRef = record.DwgRef ?? string.Empty;

                var rowUpdated = false;

                if (columnCount > 0 && ValueDiffers(rawKey, desiredCrossing)) rowUpdated = true;
                if (columnCount > 1 && ValueDiffers(ReadCellText(table, row, 1), desiredOwner)) rowUpdated = true;
                if (columnCount > 2 && ValueDiffers(ReadCellText(table, row, 2), desiredDescription)) rowUpdated = true;
                if (columnCount > 3 && ValueDiffers(ReadCellText(table, row, 3), desiredLocation)) rowUpdated = true;
                if (columnCount > 4 && ValueDiffers(ReadCellText(table, row, 4), desiredDwgRef)) rowUpdated = true;

                if (columnCount > 0) SetCellCrossingValue(table, row, 0, desiredCrossing);

                Cell ownerCell = null;
                if (columnCount > 1)
                {
                    try { ownerCell = table.Cells[row, 1]; } catch { ownerCell = null; }
                    SetCellValue(ownerCell, desiredOwner);
                }

                Cell descriptionCell = null;
                if (columnCount > 2)
                {
                    try { descriptionCell = table.Cells[row, 2]; } catch { descriptionCell = null; }
                    SetCellValue(descriptionCell, desiredDescription);
                }

                Cell locationCell = null;
                if (columnCount > 3)
                {
                    try { locationCell = table.Cells[row, 3]; } catch { locationCell = null; }
                    SetCellValue(locationCell, desiredLocation);
                }

                Cell dwgRefCell = null;
                if (columnCount > 4)
                {
                    try { dwgRefCell = table.Cells[row, 4]; } catch { dwgRefCell = null; }
                    SetCellValue(dwgRefCell, desiredDwgRef);
                }

                if (rowUpdated)
                {
                    updated++;
                    Log(string.Format(CultureInfo.InvariantCulture, "Row {0} -> UPDATED (key='{1}')", row, logKey));
                }
            }

            RefreshTable(table);
        }

        private void UpdatePageTable(Table table, IDictionary<string, CrossingRecord> byKey, out int matched, out int updated)
        {
            matched = 0;
            updated = 0;
            if (table == null) return;

            var columnCount = table.Columns.Count;
            var records = byKey?.Values ?? Enumerable.Empty<CrossingRecord>();

            for (var row = 0; row < table.Rows.Count; row++)
            {
                var rawKey = ResolveCrossingKey(table, row, 0);
                var key = NormalizeKeyForLookup(rawKey);
                CrossingRecord record = null;

                if (!string.IsNullOrEmpty(key) && byKey != null)
                {
                    if (!byKey.TryGetValue(key, out record))
                    {
                        record = byKey.Values.FirstOrDefault(r => CrossingRecord.CompareCrossingKeys(r.Crossing, key) == 0);
                    }
                }

                if (record == null)
                {
                    record = FindRecordByPageColumns(table, row, records);
                }

                if (record == null)
                {
                    Log(string.Format(CultureInfo.InvariantCulture, "Row {0} -> NO MATCH (key='{1}')", row, rawKey));
                    continue;
                }

                var logKey = !string.IsNullOrEmpty(key) ? key : (rawKey ?? string.Empty);
                matched++;

                var desiredCrossing = (record.Crossing ?? string.Empty).Trim();
                var desiredOwner = record.Owner ?? string.Empty;
                var desiredDescription = record.Description ?? string.Empty;

                var rowUpdated = false;

                if (columnCount > 0 && ValueDiffers(rawKey, desiredCrossing)) rowUpdated = true;
                if (columnCount > 1 && ValueDiffers(ReadCellText(table, row, 1), desiredOwner)) rowUpdated = true;
                if (columnCount > 2 && ValueDiffers(ReadCellText(table, row, 2), desiredDescription)) rowUpdated = true;

                if (columnCount > 0) SetCellCrossingValue(table, row, 0, desiredCrossing);

                Cell ownerCell = null;
                if (columnCount > 1)
                {
                    try { ownerCell = table.Cells[row, 1]; } catch { ownerCell = null; }
                    SetCellValue(ownerCell, desiredOwner);
                }

                Cell descriptionCell = null;
                if (columnCount > 2)
                {
                    try { descriptionCell = table.Cells[row, 2]; } catch { descriptionCell = null; }
                    SetCellValue(descriptionCell, desiredDescription);
                }

                if (rowUpdated)
                {
                    updated++;
                    Log(string.Format(CultureInfo.InvariantCulture, "Row {0} -> UPDATED (key='{1}')", row, logKey));
                }
            }

            RefreshTable(table);
        }

        private void UpdateLatLongTable(Table table, IDictionary<string, CrossingRecord> byKey, out int matched, out int updated)
        {
            matched = 0;
            updated = 0;
            if (table == null) return;

            var columnCount = table.Columns.Count;
            var records = byKey?.Values ?? Enumerable.Empty<CrossingRecord>();

            for (var row = 0; row < table.Rows.Count; row++)
            {
                var rawKey = ResolveCrossingKey(table, row, 0);
                var key = NormalizeKeyForLookup(rawKey);
                CrossingRecord record = null;

                if (!string.IsNullOrEmpty(key) && byKey != null)
                {
                    if (!byKey.TryGetValue(key, out record))
                    {
                        record = byKey.Values.FirstOrDefault(r => CrossingRecord.CompareCrossingKeys(r.Crossing, key) == 0);
                    }
                }

                if (record == null)
                {
                    record = FindRecordByLatLongColumns(table, row, records);
                }

                if (record == null)
                {
                    Log(string.Format(CultureInfo.InvariantCulture, "Row {0} -> NO MATCH (key='{1}')", row, rawKey));
                    continue;
                }

                var logKey = !string.IsNullOrEmpty(key) ? key : (rawKey ?? string.Empty);
                matched++;

                var desiredCrossing = (record.Crossing ?? string.Empty).Trim();
                var desiredDescription = record.Description ?? string.Empty;
                var desiredLat = record.Lat ?? string.Empty;
                var desiredLong = record.Long ?? string.Empty;

                var rowUpdated = false;

                if (columnCount > 0 && ValueDiffers(rawKey, desiredCrossing)) rowUpdated = true;
                if (columnCount > 1 && ValueDiffers(ReadCellText(table, row, 1), desiredDescription)) rowUpdated = true;
                if (columnCount > 2 && ValueDiffers(ReadCellText(table, row, 2), desiredLat)) rowUpdated = true;
                if (columnCount > 3 && ValueDiffers(ReadCellText(table, row, 3), desiredLong)) rowUpdated = true;

                if (columnCount > 0) SetCellCrossingValue(table, row, 0, desiredCrossing);

                Cell descriptionCell = null;
                if (columnCount > 1)
                {
                    try { descriptionCell = table.Cells[row, 1]; } catch { descriptionCell = null; }
                    SetCellValue(descriptionCell, desiredDescription);
                }

                Cell latCell = null;
                if (columnCount > 2)
                {
                    try { latCell = table.Cells[row, 2]; } catch { latCell = null; }
                    SetCellValue(latCell, desiredLat);
                }

                Cell longCell = null;
                if (columnCount > 3)
                {
                    try { longCell = table.Cells[row, 3]; } catch { longCell = null; }
                    SetCellValue(longCell, desiredLong);
                }

                if (rowUpdated)
                {
                    updated++;
                    Log(string.Format(CultureInfo.InvariantCulture, "Row {0} -> UPDATED (key='{1}')", row, logKey));
                }
            }

            RefreshTable(table);
        }

        private static bool ValueDiffers(string existing, string desired)
        {
            var left = (existing ?? string.Empty).Trim();
            var right = (desired ?? string.Empty).Trim();
            return !string.Equals(left, right, StringComparison.Ordinal);
        }

        private static string ReadCellText(Table table, int row, int column)
        {
            if (table == null || row < 0 || column < 0) return string.Empty;
            if (row >= table.Rows.Count || column >= table.Columns.Count) return string.Empty;
            try
            {
                var cell = table.Cells[row, column];
                return ReadCellText(cell);
            }
            catch { return string.Empty; }
        }

        private static bool IsCoordinateValue(string text, double min, double max)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            double value;
            if (!double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return false;
            return value >= min && value <= max;
        }

        private static string ReadCellText(Cell cell)
        {
            if (cell == null) return string.Empty;
            try { return cell.TextString ?? string.Empty; } catch { return string.Empty; }
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
                        yield return StripMTextFormatting(text).Trim();
                        continue;
                    }
                }

                // IMPORTANT: do not yield item.ToString(); it returns type names (garbage for matching).
            }
        }

        private static bool IsIntLike(Type t)
        {
            return t == typeof(short) || t == typeof(int) || t == typeof(long) ||
                   t == typeof(ushort) || t == typeof(uint) || t == typeof(ulong);
        }
        private static bool CellHasBlockContent(Cell cell)
        {
            if (cell == null) return false;
            try
            {
                var contentsProp = cell.GetType().GetProperty("Contents", BindingFlags.Public | BindingFlags.Instance);
                var contents = contentsProp?.GetValue(cell, null) as IEnumerable;
                if (contents == null) return false;

                foreach (var item in contents)
                {
                    var ctProp = item.GetType().GetProperty("ContentTypes", BindingFlags.Public | BindingFlags.Instance);
                    var ctVal = ctProp?.GetValue(item, null)?.ToString() ?? string.Empty;
                    if (ctVal.IndexOf("Block", StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
            }
            catch { }
            return false;
        }

        // NEW: Enumerate attribute TAGs from the block definition referenced by a table cell's content
        private static IEnumerable<string> GetBlockAttributeTagsFromCell(Table t, int row, int col)
        {
            if (t == null) yield break;

            Cell cell = null;
            try { cell = t.Cells[row, col]; } catch { cell = null; }
            if (cell == null) yield break;

            var contentsProp = cell.GetType().GetProperty("Contents", BindingFlags.Public | BindingFlags.Instance);
            var contents = contentsProp?.GetValue(cell, null) as IEnumerable;
            if (contents == null) yield break;

            // Use the top transaction this code already runs under (UpdateAllTables/commands open one)
            var tr = t.Database?.TransactionManager?.TopTransaction as Transaction;
            if (tr == null) yield break;

            foreach (var item in contents)
            {
                if (item == null) continue;

                var typesProp = item.GetType().GetProperty("ContentTypes", BindingFlags.Public | BindingFlags.Instance);
                var typesStr = typesProp?.GetValue(item, null)?.ToString() ?? string.Empty;
                if (typesStr.IndexOf("Block", StringComparison.OrdinalIgnoreCase) < 0) continue;

                var btrProp = item.GetType().GetProperty("BlockTableRecordId", BindingFlags.Public | BindingFlags.Instance);
                var idObj = btrProp?.GetValue(item, null);
                if (idObj is ObjectId btrId && btrId.IsValid && !btrId.IsNull)
                {
                    BlockTableRecord btr = null;
                    try { btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord; } catch { btr = null; }
                    if (btr == null) continue;

                    foreach (ObjectId entId in btr)
                    {
                        AttributeDefinition ad = null;
                        try { ad = tr.GetObject(entId, OpenMode.ForRead) as AttributeDefinition; } catch { ad = null; }
                        if (ad != null && !string.IsNullOrWhiteSpace(ad.Tag))
                            yield return ad.Tag.Trim();
                    }
                }
            }
        }


        private static string BuildHeaderLog(Table table)
        {
            if (table == null) return "headers=[]";
            var rowCount = table.Rows.Count;
            var columnCount = table.Columns.Count;
            if (rowCount == 0 || columnCount <= 0) return "headers=[]";

            var headers = new List<string>(columnCount);
            for (var col = 0; col < columnCount; col++)
            {
                string text;
                try { text = table.Cells[0, col].TextString ?? string.Empty; } catch { text = string.Empty; }
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
                var p = method.GetParameters();

                // Signature: (int row, int col, string tag, [optional ...])
                if (p.Length >= 3 &&
                    p[0].ParameterType != typeof(string) &&
                    p[1].ParameterType != typeof(string) &&
                    typeof(string).IsAssignableFrom(p[2].ParameterType))
                {
                    var args = new object[p.Length];
                    if (!TryConvertParameter(row, p[0], out args[0]) ||
                        !TryConvertParameter(col, p[1], out args[1]) ||
                        !TryConvertParameter(tag, p[2], out args[2]))
                        continue;

                    for (int i = 3; i < p.Length; i++) args[i] = p[i].IsOptional ? Type.Missing : null;
                    try
                    {
                        var result = method.Invoke(table, args);
                        var text = Convert.ToString(result, CultureInfo.InvariantCulture);
                        if (!string.IsNullOrWhiteSpace(text)) return text;
                    }
                    catch { }
                }

                // Signature: (int row, int col, int contentIndex, string tag, [optional ...])
                if (p.Length >= 4 &&
                    p[0].ParameterType != typeof(string) &&
                    p[1].ParameterType != typeof(string) &&
                    p[2].ParameterType != typeof(string) &&
                    typeof(string).IsAssignableFrom(p[3].ParameterType))
                {
                    var args = new object[p.Length];
                    if (!TryConvertParameter(row, p[0], out args[0]) ||
                        !TryConvertParameter(col, p[1], out args[1]) ||
                        !TryConvertParameter(0, p[2], out args[2]) ||
                        !TryConvertParameter(tag, p[3], out args[3]))
                        continue;

                    for (int i = 4; i < p.Length; i++) args[i] = p[i].IsOptional ? Type.Missing : null;
                    try
                    {
                        var result = method.Invoke(table, args);
                        var text = Convert.ToString(result, CultureInfo.InvariantCulture);
                        if (!string.IsNullOrWhiteSpace(text)) return text;
                    }
                    catch { }
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
                var p = method.GetParameters();

                // Signature: (int row, int col, string tag, string value, [optional ...])
                if (p.Length >= 4 &&
                    p[0].ParameterType != typeof(string) &&
                    p[1].ParameterType != typeof(string) &&
                    typeof(string).IsAssignableFrom(p[2].ParameterType) &&
                    typeof(string).IsAssignableFrom(p[3].ParameterType))
                {
                    var args = new object[p.Length];
                    if (!TryConvertParameter(row, p[0], out args[0]) ||
                        !TryConvertParameter(col, p[1], out args[1]) ||
                        !TryConvertParameter(tag, p[2], out args[2]) ||
                        !TryConvertParameter(value ?? string.Empty, p[3], out args[3]))
                        continue;

                    for (int i = 4; i < p.Length; i++) args[i] = p[i].IsOptional ? Type.Missing : null;
                    try
                    {
                        method.Invoke(table, args);
                        return true;
                    }
                    catch { }
                }

                // Signature: (int row, int col, int contentIndex, string tag, string value, [optional ...])
                if (p.Length >= 5 &&
                    p[0].ParameterType != typeof(string) &&
                    p[1].ParameterType != typeof(string) &&
                    p[2].ParameterType != typeof(string) &&
                    typeof(string).IsAssignableFrom(p[3].ParameterType) &&
                    typeof(string).IsAssignableFrom(p[4].ParameterType))
                {
                    var args = new object[p.Length];
                    if (!TryConvertParameter(row, p[0], out args[0]) ||
                        !TryConvertParameter(col, p[1], out args[1]) ||
                        !TryConvertParameter(0, p[2], out args[2]) ||
                        !TryConvertParameter(tag, p[3], out args[3]) ||
                        !TryConvertParameter(value ?? string.Empty, p[4], out args[4]))
                        continue;

                    for (int i = 5; i < p.Length; i++) args[i] = p[i].IsOptional ? Type.Missing : null;
                    try
                    {
                        method.Invoke(table, args);
                        return true;
                    }
                    catch { }
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
            if (table == null) return;

            var recompute = table.GetType().GetMethod("RecomputeTableBlock", new[] { typeof(bool) });
            if (recompute != null)
            {
                try
                {
                    recompute.Invoke(table, new object[] { true });
                    return;
                }
                catch { }
            }

            try { table.GenerateLayout(); } catch { }
        }
    }
}
