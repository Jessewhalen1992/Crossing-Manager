using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
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

        private static readonly HashSet<string> MainLocationColumnSynonyms = new HashSet<string>(StringComparer.Ordinal)
        {
            "LOCATION",
            "LOC",
            "LOCATIONS",
            "XINGLOCATION",
            "XINGLOCATIONS"
        };

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
            "XING_NO",
            "XINGNO",
            "XING#",
            "XINGNUM",
            "XING_NUM",
            "X",
            "CROSSING_ID",
            "CROSSINGID",
            "XNUM",
            "XNUMBER",
            "NUMBER",
            "INDEX",
            // Some drawings use NO as the XING bubble tag.
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

        // =====================================================================
        // Update existing tables in the drawing
        // =====================================================================

        public void UpdateAllTables(Document doc, IList<CrossingRecord> records)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (records == null) throw new ArgumentNullException(nameof(records));

            var db = doc.Database;
            _ed = doc.Editor;
            using (Logger.Scope(_ed, "update_all_tables", $"records={records.Count}"))
            {
                var byKey = records.ToDictionary(r => r.CrossingKey, r => r, StringComparer.OrdinalIgnoreCase);

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        foreach (ObjectId entId in btr)
                        {
                            var table = tr.GetObject(entId, OpenMode.ForRead) as Table;
                            if (table == null) continue;

                            var type = IdentifyTable(table, tr);
                            var typeLabel = type.ToString().ToUpperInvariant();

                            if (type == XingTableType.Unknown)
                            {
                                Logger.Info(_ed, $"table handle={entId.Handle} type={typeLabel} matched=0 updated=0 reason=unknown_type");
                                var headerLog = BuildHeaderLog(table);
                                if (!string.IsNullOrEmpty(headerLog))
                                {
                                    Logger.Debug(_ed, $"table handle={entId.Handle} header={headerLog}");
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
                                        var rowIndexMap = BuildLatLongRowIndexMap(table.ObjectId, byKey?.Values);
                                        UpdateLatLongTable(table, byKey, out matched, out updated, rowIndexMap);
                                        break;
                                }

                                _factory.TagTable(tr, table, typeLabel);
                                Logger.Info(_ed, $"table handle={entId.Handle} type={typeLabel} matched={matched} updated={updated}");
                            }
                            catch (Exception ex)
                            {
                                Logger.Warn(_ed, $"table handle={entId.Handle} type={typeLabel} err={ex.Message}");
                            }
                        }
                    }

                    tr.Commit();
                }
            }
        }



        /// <summary>
        /// Updates only Crossing tables (Main + Page) and intentionally skips LAT/LONG tables.
        /// This is used by the general Crossing duplicate resolver so it doesn't interfere with the LAT/LONG resolver.
        /// </summary>
        public void UpdateCrossingTables(Document doc, IList<CrossingRecord> records)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (records == null) throw new ArgumentNullException(nameof(records));

            var db = doc.Database;
            _ed = doc.Editor;
            using (Logger.Scope(_ed, "update_crossing_tables", $"records={records.Count}"))
            {
                var byKey = records.ToDictionary(r => r.CrossingKey, r => r, StringComparer.OrdinalIgnoreCase);

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        foreach (ObjectId entId in btr)
                        {
                            var table = tr.GetObject(entId, OpenMode.ForRead) as Table;
                            if (table == null) continue;

                            var type = IdentifyTable(table, tr);
                            if (type != XingTableType.Main && type != XingTableType.Page)
                                continue;

                            var typeLabel = type.ToString().ToUpperInvariant();

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
                                }

                                _factory.TagTable(tr, table, typeLabel);
                                Logger.Info(_ed, $"table handle={entId.Handle} type={typeLabel} matched={matched} updated={updated}");
                            }
                            catch (Exception ex)
                            {
                                Logger.Warn(_ed, $"table handle={entId.Handle} type={typeLabel} err={ex.Message}");
                            }
                        }
                    }

                    tr.Commit();
                }
            }
        }
        public void UpdateLatLongSourceTables(Document doc, IList<CrossingRecord> records)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (records == null) throw new ArgumentNullException(nameof(records));

            // Tables referenced by LAT/LONG sources discovered during scanning
            var tableIds = records
                .Where(r => r != null)
                .SelectMany(r => r.LatLongSources ?? Enumerable.Empty<CrossingRecord.LatLongSource>(),
                    (record, source) => new { Record = record, Source = source })
                .Where(x => x.Source != null && !x.Source.TableId.IsNull && x.Source.TableId.IsValid && x.Source.RowIndex >= 0)
                .Select(x => x.Source.TableId)
                .Distinct()
                .ToList();

            if (tableIds.Count == 0)
                return;

            _ed = doc.Editor;

            var byKey = records
                .Where(r => r != null)
                .ToDictionary(r => r.CrossingKey, r => r, StringComparer.OrdinalIgnoreCase);

            using (doc.LockDocument())
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                foreach (var tableId in tableIds)
                {
                    Table table = null;
                    try { table = tr.GetObject(tableId, OpenMode.ForWrite, false, true) as Table; }
                    catch { table = null; }

                    if (table == null)
                        continue;

                    var identifiedType = IdentifyTable(table, tr);
                    var rowIndexMap = BuildLatLongRowIndexMap(table.ObjectId, byKey?.Values);

                    int matched = 0, updated = 0;

                    try
                    {
                        if (identifiedType == XingTableType.LatLong)
                        {
                            // Normal header-driven updater (program-generated tables)
                            UpdateLatLongTable(table, byKey, out matched, out updated, rowIndexMap);
                        }
                        else
                        {
                            // Try normal pass anyway
                            UpdateLatLongTable(table, byKey, out matched, out updated, rowIndexMap);

                            // If nothing changed or itâ€™s clearly non-standard, try legacy row heuristics
                            if (updated == 0 || table.Columns.Count < 4)
                            {
                                updated += UpdateLegacyLatLongRows(table, rowIndexMap);
                            }
                        }

                        _factory.TagTable(tr, table, XingTableType.LatLong.ToString().ToUpperInvariant());

                        var variant = identifiedType == XingTableType.LatLong ? "standard" : "legacy";
                        Logger.Info(_ed, $"table handle={table.ObjectId.Handle} type=LATLONG variant={variant} matched={matched} updated={updated}");

                        if (updated > 0)
                        {
                            try { table.GenerateLayout(); } catch { }
                            try { table.RecordGraphicsModified(true); } catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn(_ed, $"table handle={table.ObjectId.Handle} type=LATLONG err={ex.Message}");
                    }
                }

                tr.Commit();
            }
        }

        // =====================================================================
        // PAGE TABLE CREATION (no extra title row, header bold, height=10)
        // =====================================================================

        // Returns the total width/height of a Page table for a given number of data rows.
        // Widths: 43.5, 144.5, 393.5; Row height: 25; Header rows: exactly 1.
        public void GetPageTableSize(int dataRowCount, out double totalWidth, out double totalHeight)
        {
            // Keep these in sync with CreateAndInsertPageTable
            const double W0 = 43.5;
            const double W1 = 144.5;
            const double W2 = 393.5;
            const double RowH = 25.0;
            const int HeaderRows = 1;

            totalWidth = W0 + W1 + W2;
            totalHeight = (HeaderRows + Math.Max(0, dataRowCount)) * RowH;
        }

        public void GetLatLongTableSize(int dataRowCount, out double totalWidth, out double totalHeight)
        {
            const double W0 = 40.0;   // ID
            const double W1 = 200.0;  // Description
            const double W2 = 90.0;   // Latitude
            const double W3 = 90.0;   // Longitude
            const double TitleRowHeight = 20.0;
            const double HeaderRowHeight = 25.0;
            const double DataRowHeight = 25.0;

            totalWidth = W0 + W1 + W2 + W3;
            totalHeight = TitleRowHeight + HeaderRowHeight + Math.Max(0, dataRowCount) * DataRowHeight;
        }

        public void CreateAndInsertPageTable(
            Database db,
            Transaction tr,
            BlockTableRecord ownerBtr,
            Point3d insertPoint,
            string dwgRef,
            IList<CrossingRecord> all)
        {
            // ---- fixed layout you asked for ----
            const double TextH = 10.0;
            const double RowH = 25.0;
            const double W0 = 43.5;   // XING col
            const double W1 = 144.5;  // OWNER col
            const double W2 = 393.5;  // DESCRIPTION col

            // 1) Filter + order rows for this DWG_REF
            var rows = (all ?? new List<CrossingRecord>())
                .Where(r => string.Equals(r.DwgRef ?? "", dwgRef ?? "", StringComparison.OrdinalIgnoreCase))
                .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();

            // 2) Create table, position, style
            var t = new Table();
            t.SetDatabaseDefaults();
            t.Position = insertPoint;

            var layerId = LayerUtils.EnsureLayer(db, tr, TableFactory.LayerName);
            if (!layerId.IsNull)
            {
                t.LayerId = layerId;
            }

            var tsDict = (DBDictionary)tr.GetObject(db.TableStyleDictionaryId, OpenMode.ForRead);
            if (tsDict.Contains("Standard"))
                t.TableStyle = tsDict.GetAt("Standard");

            // Build with: 1 title + 1 header + N data (we'll delete the title row to avoid the extra line)
            int headerRow = 1;
            int dataStart = 2;
            t.SetSize(2 + rows.Count, 3);

            // Column widths
            if (t.Columns.Count >= 3)
            {
                t.Columns[0].Width = W0;
                t.Columns[1].Width = W1;
                t.Columns[2].Width = W2;
            }

            // Row heights (uniform 25)
            for (int r = 0; r < t.Rows.Count; r++)
                t.Rows[r].Height = RowH;

            // 3) Header labels (no inline \b; we'll bold via a style)
            t.Cells[headerRow, 0].TextString = "XING";
            t.Cells[headerRow, 1].TextString = "OWNER";
            t.Cells[headerRow, 2].TextString = "DESCRIPTION";

            // Create/get a bold text style for headers
            var boldStyleId = EnsureBoldTextStyle(db, tr, "XING_BOLD", "Standard");

            // Apply header formatting: bold style, height, middle-center
            for (int c = 0; c < 3; c++)
            {
                var cell = t.Cells[headerRow, c];
                cell.TextHeight = TextH;
                cell.Alignment = CellAlignment.MiddleCenter;
                cell.TextStyleId = boldStyleId;
            }

            // 4) Pick a bubble block for the XING column
            var bubbleBtrId = TryFindPrototypeBubbleBlockIdFromExistingTables(db, tr);
            if (bubbleBtrId.IsNull)
                bubbleBtrId = TryFindBubbleBlockByName(db, tr, new[] { "XING_CELL", "XING_TABLE_CELL", "X_BUBBLE", "XING2", "XING" });

            // 5) Fill data rows
            for (int i = 0; i < rows.Count; i++)
            {
                var rec = rows[i];
                int row = dataStart + i;

                // Center + text height for all cells in this row
                for (int c = 0; c < 3; c++)
                {
                    var cell = t.Cells[row, c];
                    cell.Alignment = CellAlignment.MiddleCenter;
                    cell.TextHeight = TextH;
                }

                // Col 0 = bubble block (preferred) or plain text fallback
                var xVal = NormalizeXKey(rec.Crossing);
                bool placed = false;
                if (!bubbleBtrId.IsNull)
                {
                    placed = TrySetCellToBlockWithAttribute(t, row, 0, tr, bubbleBtrId, xVal);
                    if (placed)
                        TrySetCellBlockScale(t, row, 0, 1.0);   // force block scale = 1.0
                }
                if (!placed)
                    t.Cells[row, 0].TextString = xVal;

                // Col 1/2 = Owner / Description
                t.Cells[row, 1].TextString = rec.Owner ?? string.Empty;
                t.Cells[row, 2].TextString = rec.Description ?? string.Empty;
            }

            // 6) Remove the Title row so thereâ€™s no extra thin row at the top
            try { t.DeleteRows(0, 1); } catch { /* ignore if not supported */ }

            // After deletion, header is now row 0; ensure its formatting
            if (t.Rows.Count > 0)
            {
                t.Rows[0].Height = RowH;
                for (int c = 0; c < 3; c++)
                {
                    var cell = t.Cells[0, c];
                    cell.Alignment = CellAlignment.MiddleCenter;
                    cell.TextHeight = TextH;
                    cell.TextStyleId = boldStyleId;
                }
            }

            // 7) Commit entity
            ownerBtr.AppendEntity(t);
            tr.AddNewlyCreatedDBObject(t, true);
            try { t.GenerateLayout(); } catch { }
        }

        // =====================================================================
        // LAT/LONG TABLE CREATION (kept so XingForm compiles)
        // =====================================================================

        public class LatLongSection
        {
            public string Header { get; set; }

            public IList<CrossingRecord> Records { get; set; } = new List<CrossingRecord>();
        }

        private enum LatLongRowKind
        {
            Data,
            SectionHeader,
            ColumnHeader
        }

        private struct LatLongRow
        {
            private LatLongRow(LatLongRowKind kind, CrossingRecord record, string header)
            {
                Kind = kind;
                Record = record;
                Header = header;
            }

            public LatLongRowKind Kind { get; }

            public CrossingRecord Record { get; }

            public string Header { get; }

            public bool IsSectionHeader => Kind == LatLongRowKind.SectionHeader;

            public bool IsColumnHeader => Kind == LatLongRowKind.ColumnHeader;

            public bool HasRecord => Kind == LatLongRowKind.Data && Record != null;

            public static LatLongRow FromRecord(CrossingRecord record) => new LatLongRow(LatLongRowKind.Data, record, null);

            public static LatLongRow FromHeader(string header) => new LatLongRow(LatLongRowKind.SectionHeader, null, header);

            public static LatLongRow ColumnHeader() => new LatLongRow(LatLongRowKind.ColumnHeader, null, null);
        }

        public Table CreateAndInsertLatLongTable(
            Database db,
            Transaction tr,
            BlockTableRecord space,
            Point3d insertPoint,
            IList<CrossingRecord> records,
            string titleOverride = null,
            IList<LatLongSection> sections = null,
            bool includeTitleRow = true, bool includeColumnHeaders = true)
        {
            if (db == null) throw new ArgumentNullException(nameof(db));
            if (tr == null) throw new ArgumentNullException(nameof(tr));
            if (space == null) throw new ArgumentNullException(nameof(space));

            const double CellTextHeight = 10.0;
            const double TitleTextHeight = 14.0;
            const double TitleRowHeight = 20.0;
            const double HeaderRowHeight = 25.0;
            const double DataRowHeight = 25.0;
            const double W0 = 40.0;   // ID
            const double W1 = 200.0;  // Description
            const double W2 = 90.0;   // Latitude
            const double W3 = 90.0;   // Longitude

            var orderedRows = BuildLatLongRows(records, sections);
            var resolvedTitle = string.IsNullOrWhiteSpace(titleOverride)
                ? "WATER CROSSING INFORMATION"
                : titleOverride;
            var showTitle = includeTitleRow && !string.IsNullOrWhiteSpace(resolvedTitle);
            var hasSectionColumnHeaders = orderedRows.Any(r => r.IsColumnHeader);
            var showColumnHeaders = includeColumnHeaders && !hasSectionColumnHeaders;

            var table = new Table();
            table.SetDatabaseDefaults();
            table.Position = insertPoint;

            var layerId = LayerUtils.EnsureLayer(db, tr, TableFactory.LayerName);
            if (!layerId.IsNull)
                table.LayerId = layerId;

            var tsDict = (DBDictionary)tr.GetObject(db.TableStyleDictionaryId, OpenMode.ForRead);
            if (tsDict.Contains("Standard"))
                table.TableStyle = tsDict.GetAt("Standard");

            var titleRow = showTitle ? 0 : -1;
            var headerRow = showTitle ? 1 : 0;
            var dataStart = hasSectionColumnHeaders ? headerRow : headerRow + (showColumnHeaders ? 1 : 0);
            table.SetSize(dataStart + orderedRows.Count, 4);

            if (table.Columns.Count >= 4)
            {
                table.Columns[0].Width = W0;
                table.Columns[1].Width = W1;
                table.Columns[2].Width = W2;
                table.Columns[3].Width = W3;
            }

            for (int r = 0; r < table.Rows.Count; r++)
                table.Rows[r].Height = DataRowHeight;

            if (showTitle)
                table.Rows[titleRow].Height = TitleRowHeight;

            if (showColumnHeaders)
            {
                table.Rows[headerRow].Height = HeaderRowHeight;

                table.Cells[headerRow, 0].TextString = "ID";
                table.Cells[headerRow, 1].TextString = "DESCRIPTION";
                table.Cells[headerRow, 2].TextString = "LATITUDE";
                table.Cells[headerRow, 3].TextString = "LONGITUDE";
            }

            var boldStyleId = EnsureBoldTextStyle(db, tr, "XING_BOLD", "Standard");
            var headerColor = Color.FromColorIndex(ColorMethod.ByAci, 254);
            var titleColor = Color.FromColorIndex(ColorMethod.ByAci, 14);

            if (showColumnHeaders)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    var cell = table.Cells[headerRow, c];
                    cell.TextHeight = CellTextHeight;
                    cell.Alignment = CellAlignment.MiddleCenter;
                    cell.TextStyleId = boldStyleId;
                    cell.BackgroundColor = headerColor;
                }
            }

            if (showTitle)
            {
                var titleCell = table.Cells[titleRow, 0];
                table.MergeCells(CellRange.Create(table, titleRow, 0, titleRow, Math.Max(0, table.Columns.Count - 1)));
                titleCell.TextString = resolvedTitle;
                titleCell.Alignment = CellAlignment.MiddleLeft;
                titleCell.TextHeight = TitleTextHeight;
                titleCell.TextStyleId = boldStyleId;
                ApplyCellTextColor(titleCell, titleColor);
                ApplyTitleBorderStyle(table, titleRow, 0, table.Columns.Count - 1);
            }

            // Data rows
            for (int i = 0; i < orderedRows.Count; i++)
            {
                var rowInfo = orderedRows[i];
                int row = dataStart + i;

                if (rowInfo.IsSectionHeader)
                {
                    var headerText = rowInfo.Header ?? string.Empty;

                    table.Rows[row].Height = TitleRowHeight;
                    table.MergeCells(CellRange.Create(table, row, 0, row, Math.Max(0, table.Columns.Count - 1)));
                    var headerCell = table.Cells[row, 0];
                    headerCell.Alignment = CellAlignment.MiddleLeft;
                    headerCell.TextHeight = TitleTextHeight;
                    headerCell.TextStyleId = boldStyleId;
                    headerCell.TextString = headerText;
                    ApplyCellTextColor(headerCell, titleColor);
                    ApplyTitleBorderStyle(table, row, 0, table.Columns.Count - 1);
                    continue;
                }

                if (rowInfo.IsColumnHeader)
                {
                    ApplyLatLongColumnHeaderRow(table, row, HeaderRowHeight, CellTextHeight, boldStyleId, headerColor);
                    continue;
                }

                if (!rowInfo.HasRecord)
                    continue;

                var rec = rowInfo.Record;
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    var cell = table.Cells[row, c];
                    cell.Alignment = CellAlignment.MiddleCenter;
                    cell.TextHeight = CellTextHeight;
                    ApplyDataCellBorderStyle(cell);
                }

                table.Cells[row, 0].TextString = NormalizeXKey(rec?.Crossing);
                table.Cells[row, 1].TextString = rec?.Description ?? string.Empty;
                table.Cells[row, 2].TextString = rec?.Lat ?? string.Empty;
                table.Cells[row, 3].TextString = rec?.Long ?? string.Empty;

                if (i == orderedRows.Count - 1 || orderedRows[i + 1].IsSectionHeader || orderedRows[i + 1].IsColumnHeader)
                {
                    ApplyBottomBorderToRow(table, row);
                }
            }

            space.UpgradeOpen();
            space.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
            _factory.TagTable(tr, table, XingTableType.LatLong.ToString().ToUpperInvariant());
            try { table.GenerateLayout(); } catch { }
            return table;
        }

        private static List<LatLongRow> BuildLatLongRows(IList<CrossingRecord> records, IList<LatLongSection> sections)
        {
            var rows = new List<LatLongRow>();

            if (sections != null && sections.Count > 0)
            {
                foreach (var section in sections)
                {
                    if (section == null)
                        continue;

                    var header = section.Header?.Trim();
                    var sectionRecords = (section.Records ?? new List<CrossingRecord>())
                        .Where(r => r != null)
                        .ToList();

                    if (!string.IsNullOrEmpty(header) && sectionRecords.Count > 0)
                    {
                        rows.Add(LatLongRow.FromHeader(header));
                        rows.Add(LatLongRow.ColumnHeader());
                    }

                    foreach (var record in sectionRecords)
                    {
                        rows.Add(LatLongRow.FromRecord(record));
                    }
                }
            }
            else
            {
                var ordered = (records ?? new List<CrossingRecord>())
                    .Where(r => r != null)
                    .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                    .ToList();

                foreach (var record in ordered)
                {
                    rows.Add(LatLongRow.FromRecord(record));
                }
            }

            return rows;
        }

        private static void ApplyLatLongColumnHeaderRow(
            Table table,
            int row,
            double headerRowHeight,
            double cellTextHeight,
            ObjectId textStyleId,
            Color headerColor)
        {
            if (table == null)
                return;

            if (row < 0 || row >= table.Rows.Count)
                return;

            table.Rows[row].Height = headerRowHeight;

            for (int c = 0; c < table.Columns.Count; c++)
            {
                var cell = table.Cells[row, c];
                cell.Alignment = CellAlignment.MiddleCenter;
                cell.TextHeight = cellTextHeight;
                cell.TextStyleId = textStyleId;
                cell.BackgroundColor = headerColor;

                switch (c)
                {
                    case 0:
                        cell.TextString = "ID";
                        break;
                    case 1:
                        cell.TextString = "DESCRIPTION";
                        break;
                    case 2:
                        cell.TextString = "LATITUDE";
                        break;
                    case 3:
                        cell.TextString = "LONGITUDE";
                        break;
                    default:
                        cell.TextString = string.Empty;
                        break;
                }
            }
        }

        private static void ApplyTitleBorderStyle(Table table, int row, int startColumn, int endColumn)
        {
            if (table == null || table.Columns.Count == 0) return;
            if (row < 0 || row >= table.Rows.Count) return;

            startColumn = Math.Max(0, startColumn);
            endColumn = Math.Min(table.Columns.Count - 1, endColumn);

            for (int c = startColumn; c <= endColumn; c++)
            {
                ApplyTitleBorderStyle(table.Cells[row, c]);
            }
        }

        private static void ApplyTitleBorderStyle(Cell cell)
        {
            if (cell == null) return;

            try
            {
                var bordersProp = cell.GetType().GetProperty("Borders", BindingFlags.Public | BindingFlags.Instance);
                var bordersObj = bordersProp?.GetValue(cell, null);
                if (bordersObj == null) return;

                SetBorderVisible(bordersObj, "Top", false);
                SetBorderVisible(bordersObj, "Left", false);
                SetBorderVisible(bordersObj, "Right", false);
                SetBorderVisible(bordersObj, "InsideHorizontal", false);
                SetBorderVisible(bordersObj, "InsideVertical", false);
                SetBorderVisible(bordersObj, "Bottom", true);
            }
            catch
            {
                // purely cosmetic; ignore on unsupported releases
            }
        }

        private static void ApplyBottomBorderToRow(Table table, int row)
        {
            if (table == null) return;
            if (row < 0 || row >= table.Rows.Count) return;

            for (int c = 0; c < table.Columns.Count; c++)
            {
                ApplyDataCellBorderStyle(table.Cells[row, c]);
            }
        }

        private static void ApplyDataCellBorderStyle(Cell cell)
        {
            if (cell == null)
                return;

            try
            {
                var bordersProp = cell.GetType().GetProperty("Borders", BindingFlags.Public | BindingFlags.Instance);
                var bordersObj = bordersProp?.GetValue(cell, null);
                if (bordersObj == null)
                    return;

                SetBorderVisible(bordersObj, "Top", true);
                SetBorderVisible(bordersObj, "Bottom", true);
                SetBorderVisible(bordersObj, "Left", true);
                SetBorderVisible(bordersObj, "Right", true);
                SetBorderVisible(bordersObj, "InsideHorizontal", true);
                SetBorderVisible(bordersObj, "InsideVertical", true);
                SetBorderVisible(bordersObj, "Outline", true);
            }
            catch
            {
                // purely cosmetic; ignore on unsupported releases
            }
        }

        private static void ApplyCellTextColor(Cell cell, Color color)
        {
            if (cell == null) return;

            var textColorProp = cell.GetType().GetProperty("TextColor", BindingFlags.Public | BindingFlags.Instance);
            if (textColorProp != null && textColorProp.CanWrite)
            {
                try
                {
                    textColorProp.SetValue(cell, color, null);
                    return;
                }
                catch
                {
                    // ignore and fall back to ContentColor
                }
            }

            var contentColorProp = cell.GetType().GetProperty("ContentColor", BindingFlags.Public | BindingFlags.Instance);
            if (contentColorProp != null && contentColorProp.CanWrite)
            {
                try { contentColorProp.SetValue(cell, color, null); } catch { }
            }
        }

        // Portable border visibility setter across AutoCAD releases
        private static void SetBorderVisible(object borders, string memberName, bool on)
        {
            if (borders == null || string.IsNullOrEmpty(memberName)) return;

            try
            {
                // Get e.g. Top/Left/Right/Bottom/InsideHorizontal/InsideVertical/Outline
                var member = borders.GetType().GetProperty(memberName, BindingFlags.Public | BindingFlags.Instance);
                var border = member?.GetValue(borders, null);
                if (border == null) return;

                // Common property names
                var isOnProp = border.GetType().GetProperty("IsOn");
                if (isOnProp != null && isOnProp.CanWrite) { isOnProp.SetValue(border, on, null); return; }

                var visibleProp = border.GetType().GetProperty("Visible");
                if (visibleProp != null && visibleProp.CanWrite) { visibleProp.SetValue(border, on, null); return; }

                var isVisibleProp = border.GetType().GetProperty("IsVisible");
                if (isVisibleProp != null && isVisibleProp.CanWrite) { isVisibleProp.SetValue(border, on, null); return; }

                var suppressProp = border.GetType().GetProperty("Suppress");
                if (suppressProp != null && suppressProp.CanWrite) { suppressProp.SetValue(border, !on, null); return; }

                // Method alternatives
                var setOn = border.GetType().GetMethod("SetOn", new[] { typeof(bool) });
                if (setOn != null) { setOn.Invoke(border, new object[] { on }); return; }

                var setVisible = border.GetType().GetMethod("SetVisible", new[] { typeof(bool) });
                if (setVisible != null) { setVisible.Invoke(border, new object[] { on }); return; }

                var setSuppress = border.GetType().GetMethod("SetSuppress", new[] { typeof(bool) });
                if (setSuppress != null) { setSuppress.Invoke(border, new object[] { !on }); return; }
            }
            catch
            {
                // If this member isn't present in the current AutoCAD release, skip it.
            }
        }

        // TableSync.cs
        private static int UpdateLegacyLatLongRows(Table table, IDictionary<int, CrossingRecord> rowIndexMap)
        {
            if (table == null || rowIndexMap == null || rowIndexMap.Count == 0) return 0;

            int updatedCells = 0;

            foreach (var kv in rowIndexMap)
            {
                var row = kv.Key;
                var rec = kv.Value;
                if (rec == null) continue;
                if (row < 0 || row >= table.Rows.Count) continue;

                // Discover candidate cells on this row
                TryDetectLatLongCellsInRow(table, row, out int latCol, out int longCol);

                // Write LAT only if record supplies one
                if (latCol >= 0 && latCol < table.Columns.Count && !string.IsNullOrWhiteSpace(rec.Lat))
                {
                    var existing = ReadCellTextSafe(table, row, latCol);
                    var desired = rec.Lat.Trim();
                    if (!string.Equals(existing?.Trim(), desired, StringComparison.Ordinal))
                    {
                        try { table.Cells[row, latCol].TextString = desired; updatedCells++; } catch { }
                    }
                }

                // Write LONG only if record supplies one
                if (longCol >= 0 && longCol < table.Columns.Count && !string.IsNullOrWhiteSpace(rec.Long))
                {
                    var existing = ReadCellTextSafe(table, row, longCol);
                    var desired = rec.Long.Trim();
                    if (!string.Equals(existing?.Trim(), desired, StringComparison.Ordinal))
                    {
                        try { table.Cells[row, longCol].TextString = desired; updatedCells++; } catch { }
                    }
                }
            }

            return updatedCells;
        }

        // Try to find which cell(s) in a row are the Latitude and Longitude cells.
        // Strategy:
        //   1) If a cell looks like a LAT/LONG label, prefer its right neighbor as the value.
        //   2) Otherwise, pick numeric cells by range: LAT in [-90..90], LONG in [-180..180].
        //   3) If multiple candidates exist, prefer (leftmost LAT, then the nearest LONG to its right).
        // Try to find which cell(s) in a row are the Latitude and Longitude cells.
        // Strategy:
        //   1) If a cell looks like a LAT/LONG label, prefer its right neighbor as the value.
        //   2) Otherwise, pick numeric cells by range: LAT in [-90..90], LONG in [-180..180].
        //   3) If multiple candidates exist, prefer (leftmost LAT, then the nearest LONG to its right).
        // Try to find which cell(s) in a row are the Latitude and Longitude cells.
        // Strategy:
        //   1) If a cell looks like a LAT/LONG label, prefer its right neighbor as the value.
        //   2) Otherwise, pick numeric cells by range: LAT in [-90..90], LONG in [-180..180].
        //   3) If multiple candidates exist, prefer (leftmost LAT, then the nearest LONG to its right).
        private static void TryDetectLatLongCellsInRow(Table table, int row, out int latCol, out int longCol)
        {
            latCol = -1;
            longCol = -1;

            if (table == null || row < 0 || row >= table.Rows.Count)
                return;

            int cols = table.Columns.Count;
            var latLabelCols = new List<int>();
            var longLabelCols = new List<int>();
            var latNums = new List<int>();
            var longNums = new List<int>();

            // Scan the row and collect label and numeric candidates
            for (int c = 0; c < cols; c++)
            {
                var raw = ReadCellTextSafe(table, row, c);
                var txt = NormalizeText(raw); // strips mtext formatting and trims

                if (!string.IsNullOrWhiteSpace(txt))
                {
                    var up = txt.ToUpperInvariant();

                    if (up.Contains("LATITUDE") || up == "LAT" || up.StartsWith("LAT"))
                        latLabelCols.Add(c);

                    if (up.Contains("LONGITUDE") || up == "LONG" || up.StartsWith("LONG"))
                        longLabelCols.Add(c);
                }

                // Numeric candidates
                if (double.TryParse(txt, NumberStyles.Float, CultureInfo.InvariantCulture, out double val))
                {
                    if (val >= -90.0 && val <= 90.0) latNums.Add(c);
                    if (val >= -180.0 && val <= 180.0) longNums.Add(c);
                }
            }

            // Prefer explicit label -> value (right neighbor, else left neighbor)
            for (int i = 0; i < latLabelCols.Count && latCol < 0; i++)
            {
                int c = latLabelCols[i];
                int right = c + 1, left = c - 1;
                if (right < cols) latCol = right;
                else if (left >= 0) latCol = left;
            }

            for (int i = 0; i < longLabelCols.Count && longCol < 0; i++)
            {
                int c = longLabelCols[i];
                int right = c + 1, left = c - 1;
                if (right < cols) longCol = right;
                else if (left >= 0) longCol = left;
            }

            // Fill from numeric pools if still unknown
            if (latCol < 0 && latNums.Count > 0)
                latCol = latNums[0];

            if (longCol < 0 && longNums.Count > 0)
            {
                // Prefer a LONG to the right of LAT when both exist
                int pick = -1;
                for (int idx = 0; idx < longNums.Count; idx++)
                {
                    int i = longNums[idx];
                    if (i == latCol) continue;
                    if (latCol < 0 || i > latCol) { pick = i; break; }
                }
                longCol = pick >= 0 ? pick : longNums[0];
            }

            // If we accidentally picked the same column for both, try to separate
            if (latCol >= 0 && longCol == latCol && longNums.Count > 1)
            {
                int alt = -1;
                for (int k = 0; k < longNums.Count; k++)
                {
                    int i = longNums[k];
                    if (i != latCol) { alt = i; break; }
                }
                if (alt >= 0) longCol = alt;
            }
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
                                if (string.Equals(text, "LATLONG", StringComparison.OrdinalIgnoreCase))
                                {
                                    var cols = table.Columns.Count;
                                    if (cols == 4 || cols >= 6) return XingTableType.LatLong;
                                }
                            }
                        }
                    }
                }
            }

            // Header-free heuristics: MAIN=5 cols, PAGE=3 cols, LATLONG=4 cols (+ sanity)
            if (table.Columns.Count == 5) return XingTableType.Main;
            if (table.Columns.Count == 3) return XingTableType.Page;

            if ((table.Columns.Count == 4 || table.Columns.Count >= 6) && table.Rows.Count >= 1)
            {
                var headerColumns = Math.Min(table.Columns.Count, 6);
                if (HasHeaderRow(table, headerColumns, IsLatLongHeader) || LooksLikeLatLongTable(table))
                    return XingTableType.LatLong;
            }

            return XingTableType.Unknown;
        }

        // =====================================================================
        // Updaters for Main / Page / LatLong tables
        // =====================================================================

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

                if (record == null) record = FindRecordByMainColumns(table, row, records);

                if (record == null)
                {
                    Logger.Debug(_ed, $"table handle={table.ObjectId.Handle} row={row} status=no_match key={rawKey}");
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
                if (columnCount > 1) { try { ownerCell = table.Cells[row, 1]; } catch { } SetCellValue(ownerCell, desiredOwner); }

                Cell descriptionCell = null;
                if (columnCount > 2) { try { descriptionCell = table.Cells[row, 2]; } catch { } SetCellValue(descriptionCell, desiredDescription); }

                Cell locationCell = null;
                if (columnCount > 3) { try { locationCell = table.Cells[row, 3]; } catch { } SetCellValue(locationCell, desiredLocation); }

                Cell dwgRefCell = null;
                if (columnCount > 4) { try { dwgRefCell = table.Cells[row, 4]; } catch { } SetCellValue(dwgRefCell, desiredDwgRef); }

                if (rowUpdated)
                {
                    updated++;
                    Logger.Debug(_ed, $"table handle={table.ObjectId.Handle} row={row} status=updated key={logKey}");
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

                var rowOwner = ReadNorm(table, row, 1);
                var rowDesc = ReadNorm(table, row, 2);

                // Page tables (3 columns) should treat OWNER+DESCRIPTION as the row identity,
                // so swap only the X# when renumbering. Duplicate OWNER/DESCRIPTION values are
                // expected, so we accept the first match to keep X# in sync.
                if (!string.IsNullOrEmpty(rowOwner) || !string.IsNullOrEmpty(rowDesc))
                {
                    record = FindRecordByPageColumns(table, row, records);
                }

                if (record == null && !string.IsNullOrEmpty(key) && byKey != null)
                {
                    if (!byKey.TryGetValue(key, out record))
                    {
                        record = byKey.Values.FirstOrDefault(r => CrossingRecord.CompareCrossingKeys(r.Crossing, key) == 0);
                    }

                    if (record != null && (!string.IsNullOrEmpty(rowOwner) || !string.IsNullOrEmpty(rowDesc)))
                    {
                        var ownerMatches = string.IsNullOrEmpty(rowOwner) || string.Equals(Norm(record.Owner), rowOwner, StringComparison.Ordinal);
                        var descMatches = string.IsNullOrEmpty(rowDesc) || string.Equals(Norm(record.Description), rowDesc, StringComparison.Ordinal);
                        if (!ownerMatches || !descMatches)
                        {
                            record = null;
                        }
                    }
                }

                if (record == null)
                {
                    Logger.Debug(_ed, $"table handle={table.ObjectId.Handle} row={row} status=no_match key={rawKey}");
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

                // Always update the XING bubble even if OWNER/DESCRIPTION already match.
                if (columnCount > 0) SetCellCrossingValue(table, row, 0, desiredCrossing);

                Cell ownerCell = null;
                if (columnCount > 1) { try { ownerCell = table.Cells[row, 1]; } catch { } SetCellValue(ownerCell, desiredOwner); }

                Cell descriptionCell = null;
                if (columnCount > 2) { try { descriptionCell = table.Cells[row, 2]; } catch { } SetCellValue(descriptionCell, desiredDescription); }

                if (rowUpdated)
                {
                    updated++;
                    Logger.Debug(_ed, $"table handle={table.ObjectId.Handle} row={row} status=updated key={logKey}");
                }
            }

            RefreshTable(table);
        }

        // TableSync.cs
        private void UpdateLatLongTable(
            Table table,
            IDictionary<string, CrossingRecord> byKey,
            out int matched,
            out int updated,
            IDictionary<int, CrossingRecord> rowIndexMap = null)
        {
            matched = 0;
            updated = 0;
            if (table == null) return;

            var columnCount = table.Columns.Count;
            var records = byKey?.Values ?? Enumerable.Empty<CrossingRecord>();
            var recordList = records.Where(r => r != null).ToList();

            var crossDescMap = BuildCrossingDescriptionMap(recordList);
            var descriptionMap = BuildDescriptionMap(recordList);

            var dataRow = FindLatLongDataStartRow(table);
            if (dataRow < 0) dataRow = 0;

            var hasExtendedLayout = columnCount >= 6;
            var zoneColumn = hasExtendedLayout ? 2 : -1;
            var latColumn = hasExtendedLayout ? 3 : 2;
            var longColumn = hasExtendedLayout ? 4 : 3;
            var dwgColumn = hasExtendedLayout ? 5 : -1;

            for (var row = dataRow; row < table.Rows.Count; row++)
            {
                if (IsLatLongSectionHeaderRow(table, row, zoneColumn, latColumn, longColumn, dwgColumn) ||
                    IsLatLongColumnHeaderRow(table, row))
                {
                    continue;
                }

                var rawKey = ResolveCrossingKey(table, row, 0);
                var key = NormalizeKeyForLookup(rawKey);

                // IMPORTANT:
                // - For LAT/LONG tables, the row identity is the cell contents (description + coordinates),
                //   NOT the bubble value. The bubble is only used to read the existing X# for comparison/update.
                var descriptionText = ReadCellTextSafe(table, row, 1);
                var descriptionKey = NormalizeDescriptionKey(descriptionText);

                CrossingRecord record = null;

                // 1) Prefer stable row->record mapping, but only if it still matches this row's description.
                if (rowIndexMap != null && rowIndexMap.TryGetValue(row, out var mapRecord))
                {
                    if (string.IsNullOrEmpty(descriptionKey) ||
                        string.Equals(NormalizeDescriptionKey(mapRecord.Description), descriptionKey, StringComparison.Ordinal))
                    {
                        record = mapRecord;
                    }
                }

                // 2) Primary match = table cell contents (description + lat/long [+ zone/dwg])
                if (record == null)
                {
                    record = FindRecordByLatLongColumns(table, row, recordList);
                }

                // 3) If description is unique, it is safe to match on that alone.
                if (record == null && !string.IsNullOrEmpty(descriptionKey))
                {
                    if (descriptionMap.TryGetValue(descriptionKey, out var list) && list.Count == 1)
                    {
                        record = list[0];
                    }
                }

                // 4) Last resort: use existing X# from the bubble, but ONLY if it matches the row description.
                if (record == null && !string.IsNullOrEmpty(key) && byKey.TryGetValue(key, out var keyed))
                {
                    if (string.IsNullOrEmpty(descriptionKey) ||
                        string.Equals(NormalizeDescriptionKey(keyed.Description), descriptionKey, StringComparison.Ordinal))
                    {
                        record = keyed;
                    }
                }

                if (record == null)
                {
                    // Logger.Debug in this project requires (Editor, message).
                    Logger.Debug(_ed, $"Table {table.ObjectId.Handle}: no match for lat/long row {row} (x='{rawKey}', desc='{descriptionText}')");
                    continue;
                }
                var logKey = !string.IsNullOrEmpty(key) ? key : (!string.IsNullOrEmpty(rawKey) ? rawKey : (record.Crossing ?? string.Empty));
                matched++;

                var desiredCrossing = (record.Crossing ?? string.Empty).Trim();
                var desiredLat = record.Lat ?? string.Empty;
                var desiredLong = record.Long ?? string.Empty;
                var desiredZone = record.ZoneLabel ?? string.Empty;
                var desiredDwg = record.DwgRef ?? string.Empty;

                bool rowUpdated = false;

                // CROSSING is synced (ID column). DESCRIPTION is treated as the row identity and is NOT overwritten.
                if (columnCount > 0 && ValueDiffers(rawKey, desiredCrossing)) rowUpdated = true;

                if (columnCount > 0) SetCellCrossingValue(table, row, 0, desiredCrossing);
                // ZONE/LAT/LONG/DWG: preserve existing unless record provides a non-empty value
                if (zoneColumn >= 0 && !string.IsNullOrWhiteSpace(desiredZone))
                {
                    if (ValueDiffers(ReadCellText(table, row, zoneColumn), desiredZone)) rowUpdated = true;
                    Cell zoneCell = null;
                    try { zoneCell = table.Cells[row, zoneColumn]; } catch { }
                    SetCellValue(zoneCell, desiredZone);
                }

                if (latColumn >= 0 && !string.IsNullOrWhiteSpace(desiredLat))
                {
                    if (ValueDiffers(ReadCellText(table, row, latColumn), desiredLat)) rowUpdated = true;
                    Cell latCell = null;
                    try { latCell = table.Cells[row, latColumn]; } catch { }
                    SetCellValue(latCell, desiredLat);
                }

                if (longColumn >= 0 && !string.IsNullOrWhiteSpace(desiredLong))
                {
                    if (ValueDiffers(ReadCellText(table, row, longColumn), desiredLong)) rowUpdated = true;
                    Cell longCell = null;
                    try { longCell = table.Cells[row, longColumn]; } catch { }
                    SetCellValue(longCell, desiredLong);
                }

                if (dwgColumn >= 0 && !string.IsNullOrWhiteSpace(desiredDwg))
                {
                    if (ValueDiffers(ReadCellText(table, row, dwgColumn), desiredDwg)) rowUpdated = true;
                    Cell dwgCell = null;
                    try { dwgCell = table.Cells[row, dwgColumn]; } catch { }
                    SetCellValue(dwgCell, desiredDwg);
                }

                if (rowUpdated)
                {
                    updated++;
                    Logger.Debug(_ed, $"table handle={table.ObjectId.Handle} row={row} status=updated key={logKey}");
                }
            }

            RefreshTable(table);
        }

        private static bool IsLatLongSectionHeaderRow(
            Table table,
            int row,
            int zoneColumn,
            int latColumn,
            int longColumn,
            int dwgColumn)
        {
            if (table == null)
                return false;

            var crossingText = ReadCellText(table, row, 0);
            var descriptionText = ReadCellText(table, row, 1);
            var zoneText = zoneColumn >= 0 ? ReadCellText(table, row, zoneColumn) : string.Empty;
            var latText = latColumn >= 0 ? ReadCellText(table, row, latColumn) : string.Empty;
            var longText = longColumn >= 0 ? ReadCellText(table, row, longColumn) : string.Empty;
            var dwgText = dwgColumn >= 0 ? ReadCellText(table, row, dwgColumn) : string.Empty;

            if (!string.IsNullOrWhiteSpace(crossingText))
                return false;
            if (!string.IsNullOrWhiteSpace(zoneText))
                return false;
            if (!string.IsNullOrWhiteSpace(latText))
                return false;
            if (!string.IsNullOrWhiteSpace(longText))
                return false;
            if (!string.IsNullOrWhiteSpace(dwgText))
                return false;

            var descriptionKey = NormalizeDescriptionKey(descriptionText);
            if (string.IsNullOrEmpty(descriptionKey))
                return false;

            return descriptionKey.EndsWith("CROSSING INFORMATION", StringComparison.Ordinal);
        }

        private static bool IsLatLongColumnHeaderRow(Table table, int row)
        {
            if (table == null)
                return false;

            var expected = new[] { "ID", "DESCRIPTION", "LATITUDE", "LONGITUDE" };

            for (int c = 0; c < expected.Length && c < table.Columns.Count; c++)
            {
                var text = ReadCellText(table, row, c);
                if (!string.Equals(text, expected[c], StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            return true;
        }

        private static IDictionary<int, CrossingRecord> BuildLatLongRowIndexMap(
            ObjectId tableId,
            IEnumerable<CrossingRecord> records)
        {
            if (records == null)
                return new Dictionary<int, CrossingRecord>();

            return records
                .Where(r => r != null)
                .SelectMany(
                    r => (r.LatLongSources ?? new List<CrossingRecord.LatLongSource>())
                        .Where(s => s != null && !s.TableId.IsNull && s.TableId == tableId && s.RowIndex >= 0)
                        .Select(s => new { Source = s, Record = r }))
                .GroupBy(x => x.Source.RowIndex)
                .ToDictionary(g => g.Key, g => g.First().Record);
        }

        // =====================================================================
        // Small helpers
        // =====================================================================

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

        internal static string ReadCellTextSafe(Table table, int row, int column) => ReadCellText(table, row, column);

        internal static string NormalizeText(string value) => Norm(value);

        internal static int FindLatLongDataStartRow(Table table)
        {
            if (table == null)
                return 0;

            var columns = table.Columns.Count;
            if (columns >= 6)
            {
                var start = GetDataStartRow(table, 6, IsLatLongHeader);
                if (start > 0)
                    return start;
            }

            return GetDataStartRow(table, Math.Min(columns, 4), IsLatLongHeader);
        }
        // Create (or return) a bold text style. Inherits face/charset from a base style if present.
        // Create (or return) a bold text style. Portable across AutoCAD versions.
        private static ObjectId EnsureBoldTextStyle(Database db, Transaction tr, string styleName, string baseStyleName)
        {
            var tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
            if (tst.Has(styleName)) return tst[styleName];

            // Defaults
            string face = "Arial";
            bool italic = false;

            // Try to inherit face/italic from a base style if present (properties differ by version)
            if (tst.Has(baseStyleName))
            {
                try
                {
                    var baseRec = (TextStyleTableRecord)tr.GetObject(tst[baseStyleName], OpenMode.ForRead);
                    var f = baseRec.Font; // Autodesk.AutoCAD.GraphicsInterface.FontDescriptor

                    var ft = f.GetType();
                    var pTypeface = ft.GetProperty("TypeFace") ?? ft.GetProperty("Typeface"); // different versions
                    var pItalic = ft.GetProperty("Italic");

                    if (pTypeface != null)
                    {
                        var v = pTypeface.GetValue(f, null) as string;
                        if (!string.IsNullOrWhiteSpace(v)) face = v;
                    }
                    if (pItalic != null)
                    {
                        var v = pItalic.GetValue(f, null);
                        if (v is bool b) italic = b;
                    }
                }
                catch { /* fallback to defaults */ }
            }

            var rec = new TextStyleTableRecord { Name = styleName };

            // Portable ctor (charset=0, pitchAndFamily=0)
            rec.Font = new Autodesk.AutoCAD.GraphicsInterface.FontDescriptor(face, /*bold*/ true, italic, 0, 0);

            tst.UpgradeOpen();
            var id = tst.Add(rec);
            tr.AddNewlyCreatedDBObject(rec, true);
            return id;
        }

        private static bool IsCoordinateValue(string text, double min, double max)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            if (!double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var value)) return false;
            return value >= min && value <= max;
        }

        private static string ReadCellText(Cell cell)
        {
            if (cell == null) return string.Empty;
            string text = string.Empty;
            try { text = cell.TextString ?? string.Empty; } catch { text = string.Empty; }

            if (!string.IsNullOrWhiteSpace(text))
                return text;

            foreach (var run in EnumerateCellContentText(cell))
            {
                if (!string.IsNullOrWhiteSpace(run))
                    return run;
            }

            return string.Empty;
        }

        // Returns cleaned text runs from the cell's contents
        private static IEnumerable<string> EnumerateCellContentText(Cell cell)
        {
            if (cell == null) yield break;

            var enumerable = GetCellContents(cell);
            if (enumerable == null) yield break;

            foreach (var item in enumerable)
            {
                if (item == null) continue;

                var textProp = item.GetType().GetProperty("TextString", BindingFlags.Public | BindingFlags.Instance);
                if (textProp != null)
                {
                    string text = null;
                    try { text = textProp.GetValue(item, null) as string; }
                    catch { text = null; }

                    if (!string.IsNullOrWhiteSpace(text))
                        yield return StripMTextFormatting(text).Trim();
                }
            }
        }

        // Returns raw content objects (for block BTR id probing, etc.)
        private static IEnumerable<object> EnumerateCellContentObjects(Cell cell)
        {
            if (cell == null) yield break;
            var contentsProp = cell.GetType().GetProperty("Contents", BindingFlags.Public | BindingFlags.Instance);
            System.Collections.IEnumerable seq = null;
            try { seq = contentsProp?.GetValue(cell, null) as System.Collections.IEnumerable; } catch { }
            if (seq == null) yield break;
            foreach (var item in seq) if (item != null) yield return item;
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

        // ---- TABLE -> reflection helpers for block attribute value ----

        private static string TryGetBlockAttributeValue(Table table, int row, int col, string tag)
        {
            if (table == null || string.IsNullOrEmpty(tag)) return string.Empty;

            const string methodName = "GetBlockAttributeValue";
            var type = table.GetType();
            var methods = type.GetMethods().Where(m => string.Equals(m.Name, methodName, StringComparison.Ordinal));

            foreach (var method in methods)
            {
                var p = method.GetParameters();

                // (int row, int col, string tag, [optional ...])
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

                // (int row, int col, int contentIndex, string tag, [optional ...])
                if (p.Length >= 4 &&
                    p[0].ParameterType != typeof(string) &&
                    p[1].ParameterType != typeof(string) &&
                    p[2].ParameterType != typeof(string) &&
                    typeof(string).IsAssignableFrom(p[3].ParameterType))
                {
                    var anyIndex = false;
                    foreach (var contentIndex in EnumerateCellContentIndexes(table, row, col))
                    {
                        anyIndex = true;
                        var args = new object[p.Length];
                        if (!TryConvertParameter(row, p[0], out args[0]) ||
                            !TryConvertParameter(col, p[1], out args[1]) ||
                            !TryConvertParameter(contentIndex, p[2], out args[2]) ||
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

                    if (!anyIndex)
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
            }
            return string.Empty;
        }

        private static bool TrySetBlockAttributeValue(Table table, int row, int col, string tag, string value)
        {
            if (table == null || string.IsNullOrEmpty(tag)) return false;

            const string methodName = "SetBlockAttributeValue";
            var type = table.GetType();
            var methods = type.GetMethods().Where(m => string.Equals(m.Name, methodName, StringComparison.Ordinal));

            foreach (var method in methods)
            {
                var p = method.GetParameters();

                // (int row, int col, string tag, string value, [optional ...])
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
                    try { method.Invoke(table, args); return true; } catch { }
                }

                // (int row, int col, int contentIndex, string tag, string value, [optional ...])
                if (p.Length >= 5 &&
                    p[0].ParameterType != typeof(string) &&
                    p[1].ParameterType != typeof(string) &&
                    p[2].ParameterType != typeof(string) &&
                    typeof(string).IsAssignableFrom(p[3].ParameterType) &&
                    typeof(string).IsAssignableFrom(p[4].ParameterType))
                {
                    var anyIndex = false;
                    foreach (var contentIndex in EnumerateCellContentIndexes(table, row, col))
                    {
                        anyIndex = true;
                        var args = new object[p.Length];
                        if (!TryConvertParameter(row, p[0], out args[0]) ||
                            !TryConvertParameter(col, p[1], out args[1]) ||
                            !TryConvertParameter(contentIndex, p[2], out args[2]) ||
                            !TryConvertParameter(tag, p[3], out args[3]) ||
                            !TryConvertParameter(value ?? string.Empty, p[4], out args[4]))
                            continue;

                        for (int i = 5; i < p.Length; i++) args[i] = p[i].IsOptional ? Type.Missing : null;
                        try { method.Invoke(table, args); return true; } catch { }
                    }

                    if (!anyIndex)
                    {
                        var args = new object[p.Length];
                        if (!TryConvertParameter(row, p[0], out args[0]) ||
                            !TryConvertParameter(col, p[1], out args[1]) ||
                            !TryConvertParameter(0, p[2], out args[2]) ||
                            !TryConvertParameter(tag, p[3], out args[3]) ||
                            !TryConvertParameter(value ?? string.Empty, p[4], out args[4]))
                            continue;

                        for (int i = 5; i < p.Length; i++) args[i] = p[i].IsOptional ? Type.Missing : null;
                        try { method.Invoke(table, args); return true; } catch { }
                    }
                }
            }

            return false;
        }

        private static ObjectId TryGetBlockTableRecordIdFromContent(object content)
        {
            if (content == null) return ObjectId.Null;

            var prop = content.GetType().GetProperty("BlockTableRecordId");
            if (prop == null) return ObjectId.Null;

            try
            {
                if (prop.GetValue(content, null) is ObjectId id) return id;
            }
            catch { }

            return ObjectId.Null;
        }

        private static IEnumerable<int> EnumerateCellContentIndexes(Table table, int row, int col)
        {
            var cell = GetTableCell(table, row, col);
            if (cell == null) yield break;

            var enumerable = GetCellContents(cell);
            if (enumerable == null) yield break;

            var index = 0;
            foreach (var _ in enumerable) { yield return index++; }
        }

        private static string NormalizeXKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            var up = Regex.Replace(s.ToUpperInvariant(), @"\s+", "");
            var m = Regex.Match(up, @"^X0*(\d+)$");
            if (m.Success) return "X" + m.Groups[1].Value;
            m = Regex.Match(up, @"^0*(\d+)$");
            if (m.Success) return "X" + m.Groups[1].Value;
            return up;
        }

        // Try to reuse the exact bubble block used in any existing Main/Page table in the DWG
        private ObjectId TryFindPrototypeBubbleBlockIdFromExistingTables(Database db, Transaction tr)
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            foreach (ObjectId btrId in bt)
            {
                var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                foreach (ObjectId id in btr)
                {
                    if (!(tr.GetObject(id, OpenMode.ForRead) is Table table)) continue;

                    var kind = IdentifyTable(table, tr);
                    if (kind != XingTableType.Main && kind != XingTableType.Page) continue;

                    // Compute data start row by header detection (works across versions)
                    int dataRow = 0;
                    if (kind == XingTableType.Main)
                        dataRow = GetDataStartRow(table, 5, IsMainHeader);
                    else
                        dataRow = GetDataStartRow(table, 3, IsPageHeader);
                    if (dataRow <= 0) dataRow = 1; // fallback

                    if (table.Rows.Count <= dataRow || table.Columns.Count == 0) continue;

                    Cell cell = null;
                    try { cell = table.Cells[dataRow, 0]; } catch { }
                    if (cell == null) continue;

                    foreach (var content in EnumerateCellContentObjects(cell))
                    {
                        var prop = content.GetType().GetProperty("BlockTableRecordId");
                        if (prop != null)
                        {
                            try
                            {
                                var val = prop.GetValue(content, null);
                                if (val is ObjectId oid && !oid.IsNull) return oid;
                            }
                            catch { }
                        }
                    }
                }
            }
            return ObjectId.Null;
        }

        private static ObjectId TryFindBubbleBlockByName(Database db, Transaction tr, IEnumerable<string> candidateNames)
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            foreach (var name in candidateNames)
            {
                if (bt.Has(name)) return bt[name];
            }
            return ObjectId.Null;
        }

        // Put a block in a cell and set its attribute value to e.g. "X1" (supports multiple API versions)
        // Put a block in a cell and set its attribute to e.g. "X1".
        // Ensures autoscale is OFF and final scale == 1.0
        private static bool TrySetCellToBlockWithAttribute(
            Table table, int row, int col, Transaction tr, ObjectId btrId, string xValue)
        {
            if (table == null || row < 0 || col < 0 || btrId.IsNull) return false;

            Cell cell = null;
            try { cell = table.Cells[row, col]; } catch { }
            if (cell == null) return false;

            bool placed = false;

            // Try on the cell itself, then on its content items
            if (TryOnTarget(cell)) placed = true;
            if (!placed)
            {
                foreach (var content in EnumerateCellContentObjects(cell))
                {
                    if (TryOnTarget(content)) { placed = true; break; }
                }
            }

            if (placed)
            {
                // Force scale to 1 after placement (tableâ€‘level API if available, else perâ€‘content)
                TrySetCellBlockScale(table, row, col, 1.0);
            }

            return placed;

            bool TryOnTarget(object target)
            {
                if (target == null) return false;
                var typ = target.GetType();

                // Preferred API: SetBlockTableRecordId(ObjectId, bool adjustCell)
                var miSet = typ.GetMethod("SetBlockTableRecordId", new[] { typeof(ObjectId), typeof(bool) });
                if (miSet != null)
                {
                    try
                    {
                        // DO NOT autosize the cellâ€”keeps block at native scale
                        miSet.Invoke(target, new object[] { btrId, /*adjustCell*/ false });
                    }
                    catch { return false; }

                    TryDisableAutoScale(target);                  // make sure autoscale is OFF
                    return TrySetAttribute(target, btrId, xValue, tr);
                }

                // Fallback: property BlockTableRecordId
                var pi = typ.GetProperty("BlockTableRecordId");
                if (pi != null && pi.CanWrite)
                {
                    try { pi.SetValue(target, btrId, null); } catch { return false; }
                    TryDisableAutoScale(target);
                    return TrySetAttribute(target, btrId, xValue, tr);
                }

                return false;
            }

            bool TrySetAttribute(object target, ObjectId bid, string val, Transaction t)
            {
                // By tag
                var miByTag = target.GetType().GetMethod("SetBlockAttributeValue", new[] { typeof(string), typeof(string) });
                if (miByTag != null)
                {
                    foreach (var tag in CrossingAttributeTags)
                    {
                        try { miByTag.Invoke(target, new object[] { tag, val }); return true; } catch { }
                    }
                }

                // By AttributeDefinition id
                var miById = target.GetType().GetMethod("SetBlockAttributeValue", new[] { typeof(ObjectId), typeof(string) });
                if (miById != null)
                {
                    var btr = (BlockTableRecord)t.GetObject(bid, OpenMode.ForRead);
                    foreach (ObjectId id in btr)
                    {
                        var ad = t.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                        if (ad == null) continue;
                        try { miById.Invoke(target, new object[] { ad.ObjectId, val }); return true; } catch { }
                    }
                }

                return false;
            }
        }

        // Turn OFF autoâ€‘scale/autoâ€‘fit flags that shrink the block (varies by release)
        private static void TryDisableAutoScale(object target)
        {
            if (target == null) return;

            // Properties weâ€™ve seen across versions
            var pAuto = target.GetType().GetProperty("AutoScale")
                     ?? target.GetType().GetProperty("IsAutoScale")
                     ?? target.GetType().GetProperty("AutoFit");
            if (pAuto != null && pAuto.CanWrite) { try { pAuto.SetValue(target, false, null); } catch { } }

            // Methods weâ€™ve seen across versions
            var mAuto = target.GetType().GetMethod("SetAutoScale", new[] { typeof(bool) })
                     ?? target.GetType().GetMethod("SetAutoFit", new[] { typeof(bool) });
            if (mAuto != null) { try { mAuto.Invoke(target, new object[] { false }); } catch { } }
        }

        // Try to set the block scale in the cell to a fixed value (tableâ€‘level API first)
        private static bool TrySetScaleOnTarget(object target, double sc)
        {
            if (target == null) return false;

            // Common property names for a single uniform scale
            var p = target.GetType().GetProperty("BlockScale") ?? target.GetType().GetProperty("Scale");
            if (p != null && p.CanWrite)
            {
                try { p.SetValue(target, sc, null); return true; } catch { }
            }

            // Try per-axis scale properties
            var px = target.GetType().GetProperty("ScaleX") ?? target.GetType().GetProperty("XScale");
            var py = target.GetType().GetProperty("ScaleY") ?? target.GetType().GetProperty("YScale");
            var pz = target.GetType().GetProperty("ScaleZ") ?? target.GetType().GetProperty("ZScale");

            bool ok = false;
            if (px != null && px.CanWrite) { try { px.SetValue(target, sc, null); ok = true; } catch { } }
            if (py != null && py.CanWrite) { try { py.SetValue(target, sc, null); ok = true; } catch { } }
            if (pz != null && pz.CanWrite) { try { pz.SetValue(target, sc, null); ok = true; } catch { } }

            return ok;
        }

        // Prefer table-level SetBlockTableRecordId if the API exposes it (not required here, but handy)
        private static bool TrySetCellBlockTableRecordId(Table table, int row, int col, ObjectId btrId)
        {
            if (table == null || btrId.IsNull) return false;

            foreach (var m in table.GetType().GetMethods().Where(x => x.Name == "SetBlockTableRecordId"))
            {
                var p = m.GetParameters();

                // (int row, int col, ObjectId btr, bool adjustCell)
                if (p.Length == 4 &&
                    p[0].ParameterType == typeof(int) &&
                    p[1].ParameterType == typeof(int) &&
                    p[2].ParameterType == typeof(ObjectId) &&
                    p[3].ParameterType == typeof(bool))
                {
                    try { m.Invoke(table, new object[] { row, col, btrId, true }); return true; } catch { }
                }

                // (int row, int col, int contentIndex, ObjectId btr, bool adjustCell)
                if (p.Length == 5 &&
                    p[0].ParameterType == typeof(int) &&
                    p[1].ParameterType == typeof(int) &&
                    p[2].ParameterType == typeof(int) &&
                    p[3].ParameterType == typeof(ObjectId) &&
                    p[4].ParameterType == typeof(bool))
                {
                    try { m.Invoke(table, new object[] { row, col, 0, btrId, true }); return true; } catch { }
                }
            }

            return false;
        }

        private static Cell GetTableCell(Table table, int row, int col)
        {
            if (table == null) return null;

            try { return table.Cells[row, col]; }
            catch { return null; }
        }

        private static IEnumerable GetCellContents(Cell cell)
        {
            if (cell == null) return null;

            var contentsProperty = cell.GetType().GetProperty("Contents", BindingFlags.Public | BindingFlags.Instance);
            if (contentsProperty == null) return null;

            try { return contentsProperty.GetValue(cell, null) as IEnumerable; }
            catch { return null; }
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

        // ---------- header detection helpers ----------

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
            if (headers == null)
                return false;

            var extended = new[] { "ID", "DESCRIPTION", "ZONE", "LATITUDE", "LONGITUDE", "DWG_REF" };
            if (CompareHeaders(headers, extended)) return true;

            var updated = new[] { "ID", "DESCRIPTION", "LATITUDE", "LONGITUDE" };
            if (CompareHeaders(headers, updated)) return true;

            var legacy = new[] { "XING", "DESCRIPTION", "LAT", "LONG" };
            return CompareHeaders(headers, legacy);
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

        internal static bool HasMainHeaderRow(Table table)
        {
            if (table == null || table.Columns.Count < 5) return false;
            return HasHeaderRow(table, 5, IsMainHeader);
        }

        internal static bool HasPageHeaderRow(Table table)
        {
            if (table == null || table.Columns.Count < 3) return false;
            return HasHeaderRow(table, 3, IsPageHeader);
        }

        internal static bool HasLatLongHeaderRow(Table table)
        {
            if (table == null) return false;

            if (table.Columns.Count >= 6 && HasHeaderRow(table, 6, IsLatLongHeader))
                return true;

            if (table.Columns.Count >= 4 && HasHeaderRow(table, 4, IsLatLongHeader))
                return true;

            return false;
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
                if (char.IsWhiteSpace(ch) || ch == '.' || ch == ',' || ch == '#' || ch == '_' ||
                    ch == '-' || ch == '/' || ch == '(' || ch == ')' || ch == '%' || ch == '|')
                    continue;
                builder.Append(char.ToUpperInvariant(ch));
            }

            var normalized = builder.ToString();
            if (columnIndex == 3 && MainLocationColumnSynonyms.Contains(normalized))
                return "LOCATION";

            if (columnIndex == 4 || columnIndex == 5)
            {
                if (MainDwgColumnSynonyms.Contains(normalized)) return "DWGREF";
            }
            return normalized;
        }

        private static readonly Regex InlineFormatRegex = new Regex(@"\\[^;\\{}]*;", RegexOptions.Compiled);
        private static readonly Regex ResidualFormatRegex = new Regex(@"\\[^{}]", RegexOptions.Compiled);
        private static readonly Regex SpecialCodeRegex = new Regex("%%[^\\s]+", RegexOptions.Compiled);

        // Strip inline MTEXT formatting commands (\H, \f, \C, \A, \S etc.)
        private static string StripMTextFormatting(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            var withoutCommands = InlineFormatRegex.Replace(value, string.Empty);
            var withoutResidual = ResidualFormatRegex.Replace(withoutCommands, string.Empty);
            var withoutSpecial = SpecialCodeRegex.Replace(withoutResidual, string.Empty);
            return withoutSpecial.Replace("{", string.Empty).Replace("}", string.Empty);
        }

        // ---------- Matching helpers used by updaters ----------

        private static bool LooksLikeLatLongTable(Table table)
        {
            if (table == null) return false;

            var rowCount = table.Rows.Count;
            if (rowCount <= 0) return false;

            var columns = table.Columns.Count;
            if (columns != 4 && columns != 6) return false;

            var rowsToScan = Math.Min(rowCount, MaxHeaderRowsToScan);
            var candidates = 0;

            for (var row = 0; row < rowsToScan; row++)
            {
                var latColumn = columns >= 6 ? 3 : 2;
                var longColumn = columns >= 6 ? 4 : 3;
                var latText = ReadCellText(table, row, latColumn);
                var longText = ReadCellText(table, row, longColumn);

                if (IsCoordinateValue(latText, -90.0, 90.0) && IsCoordinateValue(longText, -180.0, 180.0))
                {
                    var crossing = ResolveCrossingKey(table, row, 0);
                    if (string.IsNullOrWhiteSpace(crossing))
                        continue;
                    candidates++;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(latText) && string.IsNullOrWhiteSpace(longText))
                    continue;

                var normalizedLat = NormalizeHeader(latText, latColumn);
                var normalizedLong = NormalizeHeader(longText, longColumn);
                if (normalizedLat.StartsWith("LAT", StringComparison.Ordinal) &&
                    normalizedLong.StartsWith("LONG", StringComparison.Ordinal))
                {
                    continue;
                }

                return false;
            }

            return candidates > 0;
        }

        private static void SetCellCrossingValue(Table t, int row, int col, string crossingText)
        {
            if (t == null) return;

            // 1) Try to set the block attribute by any of our known tags
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
                if (TrySetBlockAttributeValueOnContent(t, cell, crossingText)) return;
                return;
            }

            // 3) Fallback to plain text when no block
            try
            {
                if (cell != null) cell.TextString = crossingText ?? string.Empty;
            }
            catch { }
        }

        private static bool TrySetBlockAttributeValueOnContent(Table table, Cell cell, string value)
        {
            if (table == null || cell == null) return false;

            var tr = table.Database?.TransactionManager?.TopTransaction as Transaction;
            foreach (var content in EnumerateCellContentObjects(cell))
            {
                if (TrySetBlockAttributeValueOnTarget(content, value, tr)) return true;
            }

            return false;
        }

        private static bool TrySetBlockAttributeValueOnTarget(object target, string value, Transaction tr)
        {
            if (target == null) return false;

            var miByTag = target.GetType().GetMethod("SetBlockAttributeValue", new[] { typeof(string), typeof(string) });
            if (miByTag != null)
            {
                foreach (var tag in CrossingAttributeTags)
                {
                    try
                    {
                        miByTag.Invoke(target, new object[] { tag, value ?? string.Empty });
                        return true;
                    }
                    catch { }
                }
            }

            var miById = target.GetType().GetMethod("SetBlockAttributeValue", new[] { typeof(ObjectId), typeof(string) });
            if (miById == null || tr == null) return false;

            var btrId = TryGetBlockTableRecordIdFromContent(target);
            if (btrId == ObjectId.Null || !btrId.IsValid) return false;

            var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
            if (btr == null) return false;

            foreach (ObjectId id in btr)
            {
                var ad = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                if (ad == null) continue;

                try
                {
                    miById.Invoke(target, new object[] { ad.ObjectId, value ?? string.Empty });
                    return true;
                }
                catch { }
            }

            return false;
        }

        private static void SetCellValue(Cell cell, string value)
        {
            if (cell == null) return;

            var desired = value ?? string.Empty;

            try
            {
                cell.TextString = desired;
            }
            catch
            {
                try
                {
                    cell.Value = desired;
                }
                catch
                {
                    // best effort: ignore assignment failures
                }
            }
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

            // If the cell contains a block (the bubble), don't trust TextString as the crossing key.
            // We only use the bubble to READ the X# via its attribute.
            bool hasBlockContent = false;
            try { hasBlockContent = cell != null && CellHasBlockContent(cell); } catch { hasBlockContent = false; }

            if (!string.IsNullOrWhiteSpace(cleanedDirect) && !hasBlockContent) return cleanedDirect;

            foreach (var tag in CrossingAttributeTags)
            {
                var blockVal = TableCellProbe.TryGetCellBlockAttr(table, row, col, tag);
                var cleaned = CleanCellText(blockVal);
                if (!string.IsNullOrWhiteSpace(cleaned)) return cleaned;
            }

            foreach (var tag in CrossingAttributeTags)
            {
                var blockValue = TryGetBlockAttributeValue(table, row, col, tag);
                var cleanedBlockText = CleanCellText(blockValue);
                if (!string.IsNullOrWhiteSpace(cleanedBlockText)) return cleanedBlockText;
            }

            // Fallback: if TextString looks like an X#, accept it (some tables use plain text instead of a bubble block).
            if (!string.IsNullOrWhiteSpace(cleanedDirect))
            {
                var norm = NormalizeCrossingForMap(cleanedDirect);
                if (!string.IsNullOrEmpty(norm)) return cleanedDirect;
            }

            var attrProperty = cell?.GetType().GetProperty("BlockAttributeValue", BindingFlags.Public | BindingFlags.Instance);
            if (attrProperty != null)
            {
                try
                {
                    var attrValue = attrProperty.GetValue(cell, null) as string;
                    var cleanedAttrText = CleanCellText(attrValue);
                    if (!string.IsNullOrWhiteSpace(cleanedAttrText)) return cleanedAttrText;
                }
                catch { }
            }

            foreach (var textRun in EnumerateCellContentText(cell))
            {
                var cleanedContent = CleanCellText(textRun);
                if (!string.IsNullOrWhiteSpace(cleanedContent)) return cleanedContent;
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

        private static string ReadNorm(Table t, int row, int col) => Norm(ReadCellText(t, row, col));
        // Try to set the scale of a block hosted in a table cell to a specific value (e.g., 1.0)
        private static bool TrySetCellBlockScale(Table table, int row, int col, double scale)
        {
            if (table == null) return false;

            // Preferred: table-level API (varies by release)
            foreach (var m in table.GetType().GetMethods().Where(x => x.Name == "SetBlockScale"))
            {
                var p = m.GetParameters();

                // (int row, int col, double scale)
                if (p.Length == 3 &&
                    p[0].ParameterType == typeof(int) &&
                    p[1].ParameterType == typeof(int) &&
                    p[2].ParameterType == typeof(double))
                {
                    try { m.Invoke(table, new object[] { row, col, scale }); return true; } catch { }
                }

                // (int row, int col, int contentIndex, double scale)
                if (p.Length == 4 &&
                    p[0].ParameterType == typeof(int) &&
                    p[1].ParameterType == typeof(int) &&
                    p[2].ParameterType == typeof(int) &&
                    p[3].ParameterType == typeof(double))
                {
                    try { m.Invoke(table, new object[] { row, col, 0, scale }); return true; } catch { }
                }
            }

            // Fallback: set on the cell or its content objects
            Cell cell = null;
            try { cell = table.Cells[row, col]; } catch { }
            if (cell != null)
            {
                if (TrySetScaleOnTarget(cell, scale)) return true;

                // Enumerate cell contents (you already have EnumerateCellContentObjects in this file)
                foreach (var item in EnumerateCellContentObjects(cell))
                    if (TrySetScaleOnTarget(item, scale)) return true;
            }

            return false;
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
            if (candidates.Count > 0) return candidates[0];

            candidates = records.Where(r => string.Equals(Norm(r.Description), desc, StringComparison.Ordinal)).ToList();
            return candidates.Count > 0 ? candidates[0] : null;
        }

        private static CrossingRecord FindRecordByLatLongColumns(Table t, int row, IEnumerable<CrossingRecord> records)
        {
            var desc = ReadNorm(t, row, 1);
            var columns = t?.Columns.Count ?? 0;
            var zone = columns >= 6 ? ReadNorm(t, row, 2) : string.Empty;
            var lat = ReadNorm(t, row, columns >= 6 ? 3 : 2);
            var lng = ReadNorm(t, row, columns >= 6 ? 4 : 3);
            var dwg = columns >= 6 ? ReadNorm(t, row, 5) : string.Empty;

            var candidates = records.Where(r =>
                string.Equals(Norm(r.Description), desc, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Lat), lat, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Long), lng, StringComparison.Ordinal) &&
                (string.IsNullOrEmpty(zone) || string.Equals(Norm(r.Zone), zone, StringComparison.Ordinal)) &&
                (string.IsNullOrEmpty(dwg) || string.Equals(Norm(r.DwgRef), dwg, StringComparison.Ordinal))).ToList();
            if (candidates.Count == 1) return candidates[0];

            candidates = records.Where(r => string.Equals(Norm(r.Description), desc, StringComparison.Ordinal)).ToList();
            return candidates.Count == 1 ? candidates[0] : null;
        }

        internal static string NormalizeKeyForLookup(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return string.Empty;

            var cleaned = StripMTextFormatting(s).Trim();
            if (string.IsNullOrEmpty(cleaned))
                return string.Empty;

            var builder = new StringBuilder(cleaned.Length);
            var previousWhitespace = false;

            foreach (var ch in cleaned)
            {
                if (char.IsWhiteSpace(ch))
                {
                    if (!previousWhitespace)
                    {
                        builder.Append(' ');
                        previousWhitespace = true;
                    }
                    continue;
                }

                previousWhitespace = false;
                builder.Append(char.IsLetter(ch) ? char.ToUpperInvariant(ch) : ch);
            }

            return builder.ToString().Trim();
        }

        private static Dictionary<string, CrossingRecord> BuildCrossingDescriptionMap(IEnumerable<CrossingRecord> records)
        {
            var map = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            if (records == null) return map;

            foreach (var record in records)
            {
                if (record == null) continue;

                var descKey = NormalizeDescriptionKey(record.Description);
                if (string.IsNullOrEmpty(descKey)) continue;

                AddCrossingDescriptionKey(map, record.CrossingKey, descKey, record);
                AddCrossingDescriptionKey(map, NormalizeKeyForLookup(record.Crossing), descKey, record);
                AddCrossingDescriptionKey(map, NormalizeCrossingForMap(record.Crossing), descKey, record);
            }

            return map;
        }

        private static Dictionary<string, List<CrossingRecord>> BuildDescriptionMap(IEnumerable<CrossingRecord> records)
        {
            var map = new Dictionary<string, List<CrossingRecord>>(StringComparer.OrdinalIgnoreCase);
            if (records == null) return map;

            foreach (var record in records)
            {
                if (record == null) continue;

                var descKey = NormalizeDescriptionKey(record.Description);
                if (string.IsNullOrEmpty(descKey)) continue;

                if (!map.TryGetValue(descKey, out var list))
                {
                    list = new List<CrossingRecord>();
                    map[descKey] = list;
                }

                list.Add(record);
            }

            return map;
        }

        private static void AddCrossingDescriptionKey(
            IDictionary<string, CrossingRecord> map,
            string crossingKey,
            string descriptionKey,
            CrossingRecord record)
        {
            if (map == null || record == null) return;
            var normalizedCrossing = NormalizeCrossingForMap(crossingKey);
            if (string.IsNullOrEmpty(normalizedCrossing) || string.IsNullOrEmpty(descriptionKey)) return;

            var composite = BuildCrossingDescriptionKey(normalizedCrossing, descriptionKey);
            if (string.IsNullOrEmpty(composite)) return;

            if (!map.ContainsKey(composite))
                map[composite] = record;
        }

        private static CrossingRecord TryResolveByCrossingAndDescription(
            string key,
            string normalizedCrossing,
            string descriptionKey,
            IDictionary<string, CrossingRecord> map)
        {
            if (map == null || string.IsNullOrEmpty(descriptionKey))
                return null;

            var searchKeys = new List<string>
            {
                BuildCrossingDescriptionKey(normalizedCrossing, descriptionKey),
                BuildCrossingDescriptionKey(key, descriptionKey)
            };

            foreach (var searchKey in searchKeys)
            {
                if (string.IsNullOrEmpty(searchKey))
                    continue;

                if (map.TryGetValue(searchKey, out var match) && match != null)
                    return match;
            }

            return null;
        }

        private static bool MatchesCrossingAndDescription(
            CrossingRecord record,
            string key,
            string normalizedCrossing,
            string descriptionKey)
        {
            if (record == null || string.IsNullOrEmpty(descriptionKey))
                return false;

            var recordDescription = NormalizeDescriptionKey(record.Description);
            if (!string.Equals(recordDescription, descriptionKey, StringComparison.OrdinalIgnoreCase))
                return false;

            var recordKey = BuildCrossingDescriptionKey(NormalizeCrossingForMap(record.Crossing), descriptionKey);
            if (string.Equals(recordKey, BuildCrossingDescriptionKey(normalizedCrossing, descriptionKey), StringComparison.OrdinalIgnoreCase))
                return true;

            if (!string.IsNullOrEmpty(key))
            {
                var alt = BuildCrossingDescriptionKey(NormalizeCrossingForMap(key), descriptionKey);
                if (string.Equals(recordKey, alt, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        private static string NormalizeCrossingForMap(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;

            var cleaned = StripMTextFormatting(value).Trim().ToUpperInvariant();
            var sb = new StringBuilder(cleaned.Length);
            foreach (var ch in cleaned)
            {
                if (!char.IsWhiteSpace(ch))
                    sb.Append(ch);
            }

            return sb.ToString();
        }

        private static string NormalizeDescriptionKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            return StripMTextFormatting(value).Trim().ToUpperInvariant();
        }

        private static string BuildCrossingDescriptionKey(string crossingKey, string descriptionKey)
        {
            if (string.IsNullOrEmpty(crossingKey) || string.IsNullOrEmpty(descriptionKey))
                return string.Empty;

            return crossingKey + "\u001F" + descriptionKey;
        }
    }
}
