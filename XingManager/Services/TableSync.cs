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
        private const double CrossingBubbleScale = 0.980;

        private static readonly Regex MTextFormattingCommandRegex = new Regex("\\\\[^;\\\\{}]*;", RegexOptions.Compiled);
        private static readonly Regex MTextResidualCommandRegex = new Regex("\\\\[^{}]", RegexOptions.Compiled);
        private static readonly Regex MTextSpecialCodeRegex = new Regex("%%[^\\s]+", RegexOptions.Compiled);
        private static readonly Regex LayoutDwgRefRegex = new Regex("^\\s*(?:XING|CROSSING)(?:\\s*DWG)?\\s*#\\s*(?<ref>.+?)\\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex LayoutCloneSuffixRegex = new Regex("^(?<base>.+)-\\d+$", RegexOptions.Compiled);
        private static readonly Regex MeridianLocationRegex = new Regex(
            "(?:N\\.?\\s*[EW]\\.?|S\\.?\\s*[EW]\\.?)" +
            "(?:\\s+|\\\\P|\\\\~)*1/4(?:\\s+|\\\\P|\\\\~)*SEC\\." +
            "(?:\\s+|\\\\P|\\\\~)*\\d+\\s*,\\s*TWP\\.(?:\\s+|\\\\P|\\\\~)*\\d+" +
            "\\s*,\\s*RGE\\.(?:\\s+|\\\\P|\\\\~)*\\d+\\s*,\\s*W\\." +
            "(?:\\s+|\\\\P|\\\\~)*\\d+(?:\\s+|\\\\P|\\\\~)*M\\.",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly LayoutUtils LocationFormatter = new LayoutUtils();

        private sealed class TableLayoutContext
        {
            public string LayoutName { get; set; } = string.Empty;

            public string DwgRef { get; set; } = string.Empty;

            public HashSet<string> LocationKeys { get; } = new HashSet<string>(StringComparer.Ordinal);

            public bool HasFilters => !string.IsNullOrWhiteSpace(DwgRef) || LocationKeys.Count > 0;
        }

        private sealed class WaterTableCreatePlan
        {
            public ObjectId LayoutBtrId { get; set; } = ObjectId.Null;

            public string LayoutName { get; set; } = string.Empty;

            public Point3d InsertPoint { get; set; } = Point3d.Origin;

            public List<CrossingRecord> Records { get; set; } = new List<CrossingRecord>();
        }

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
                                        var pageRowIndexMap = BuildPageRowIndexMap(table.ObjectId, byKey?.Values);
                                        var pageContext = BuildTableLayoutContext(table);
                                        UpdatePageTable(table, byKey, out matched, out updated, pageRowIndexMap, pageContext);
                                        break;
                                    case XingTableType.LatLong:
                                        var rowIndexMap = BuildLatLongRowIndexMap(table.ObjectId, byKey?.Values);
                                        var latLongContext = BuildTableLayoutContext(table);
                                        UpdateLatLongTable(table, byKey, out matched, out updated, rowIndexMap, latLongContext);
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

                    // Ensure water-coordinate crossings always have a WATER table on crossing layouts.
                    // This avoids a water crossing being represented in PAGE without its coordinate table.
                    EnsureWaterLatLongTablesOnCrossingLayouts(db, tr, records);

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
                                        var pageRowIndexMap = BuildPageRowIndexMap(table.ObjectId, byKey?.Values);
                                        var pageContext = BuildTableLayoutContext(table);
                                        UpdatePageTable(table, byKey, out matched, out updated, pageRowIndexMap, pageContext);
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

        private void EnsureWaterLatLongTablesOnCrossingLayouts(Database db, Transaction tr, IEnumerable<CrossingRecord> records)
        {
            if (db == null || tr == null)
                return;

            var normalizedWaterRecords = (records ?? Enumerable.Empty<CrossingRecord>())
                .Where(r => r != null)
                .Where(IsWaterCoordinateCrossingForPage)
                .Where(r => !IsPlaceholderDwgRef(r.DwgRef))
                .GroupBy(r => NormalizeDwgRefForLookup(r.DwgRef), StringComparer.OrdinalIgnoreCase)
                .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing)).ToList(),
                    StringComparer.OrdinalIgnoreCase);

            if (normalizedWaterRecords.Count == 0)
                return;

            var plans = new List<WaterTableCreatePlan>();
            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            foreach (ObjectId btrId in blockTable)
            {
                var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                if (btr == null || btr.LayoutId.IsNull)
                    continue;

                Layout layout = null;
                try { layout = tr.GetObject(btr.LayoutId, OpenMode.ForRead, false) as Layout; }
                catch { layout = null; }
                if (layout == null)
                    continue;

                var layoutRef = TryExtractDwgRefFromLayoutName(layout.LayoutName);
                if (string.IsNullOrWhiteSpace(layoutRef))
                    continue;

                if (!normalizedWaterRecords.TryGetValue(layoutRef, out var waterRecords) || waterRecords.Count == 0)
                    continue;

                var hasWaterTable = false;
                Table anchorPageTable = null;
                var waterKeys = new HashSet<string>(
                    waterRecords
                        .Select(r => NormalizeCrossingForMap(r?.Crossing))
                        .Where(k => !string.IsNullOrWhiteSpace(k)),
                    StringComparer.OrdinalIgnoreCase);

                foreach (ObjectId entId in btr)
                {
                    if (!entId.ObjectClass.DxfName.Equals("ACAD_TABLE", StringComparison.OrdinalIgnoreCase))
                        continue;

                    Table table = null;
                    try { table = tr.GetObject(entId, OpenMode.ForRead, false) as Table; }
                    catch { table = null; }
                    if (table == null)
                        continue;

                    var kind = XingTableType.Unknown;
                    try { kind = IdentifyTable(table, tr); }
                    catch { kind = XingTableType.Unknown; }

                    if (kind == XingTableType.Page && anchorPageTable == null)
                        anchorPageTable = table;

                    // Existing water table detection:
                    // 1) explicit WATER title; or
                    // 2) LAT/LONG table that already contains any water crossing key.
                    if (IsWaterCrossingLatLongTable(table))
                    {
                        hasWaterTable = true;
                        break;
                    }

                    if (kind == XingTableType.LatLong && TableContainsAnyCrossingKeys(table, waterKeys))
                    {
                        hasWaterTable = true;
                        break;
                    }
                }

                if (hasWaterTable || anchorPageTable == null)
                {
                    if (hasWaterTable)
                        Logger.Debug(_ed, $"layout={layout.LayoutName} action=create_water_latlong status=skip reason=existing_table");
                    continue;
                }

                var insertPoint = ResolveWaterTableInsertPoint(anchorPageTable, waterRecords.Count);
                plans.Add(new WaterTableCreatePlan
                {
                    LayoutBtrId = btr.ObjectId,
                    LayoutName = layout.LayoutName ?? string.Empty,
                    InsertPoint = insertPoint,
                    Records = waterRecords
                });
            }

            foreach (var plan in plans)
            {
                if (plan == null || plan.LayoutBtrId.IsNull || plan.Records == null || plan.Records.Count == 0)
                    continue;

                try
                {
                    var btr = tr.GetObject(plan.LayoutBtrId, OpenMode.ForWrite, false) as BlockTableRecord;
                    if (btr == null)
                        continue;

                    CreateAndInsertLatLongTable(db, tr, btr, plan.InsertPoint, plan.Records);
                    Logger.Info(_ed, $"layout={plan.LayoutName} action=create_water_latlong status=ok rows={plan.Records.Count}");
                }
                catch (Exception ex)
                {
                    Logger.Warn(_ed, $"layout={plan.LayoutName} action=create_water_latlong status=err err={ex.Message}");
                }
            }
        }

        private Point3d ResolveWaterTableInsertPoint(Table pageTable, int recordCount)
        {
            if (pageTable == null)
                return Point3d.Origin;

            Point3d basePoint;
            try { basePoint = pageTable.Position; }
            catch { basePoint = Point3d.Origin; }

            var pageWidth = GetTableTotalWidth(pageTable);
            var pageHeight = GetTableTotalHeight(pageTable);

            GetLatLongTableSize(Math.Max(0, recordCount), out var waterWidth, out _);

            const double verticalGap = 10.0;
            var x = basePoint.X + ((pageWidth - waterWidth) / 2.0);
            var y = basePoint.Y + pageHeight + verticalGap;

            return new Point3d(x, y, basePoint.Z);
        }

        private static double GetTableTotalWidth(Table table)
        {
            if (table == null)
                return 0.0;

            var total = 0.0;
            try
            {
                for (var i = 0; i < table.Columns.Count; i++)
                    total += table.Columns[i].Width;
            }
            catch { }

            return total;
        }

        private static double GetTableTotalHeight(Table table)
        {
            if (table == null)
                return 0.0;

            var total = 0.0;
            try
            {
                for (var i = 0; i < table.Rows.Count; i++)
                    total += table.Rows[i].Height;
            }
            catch { }

            return total;
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
                    var latLongContext = BuildTableLayoutContext(table);

                    int matched = 0, updated = 0;

                    try
                    {
                        if (identifiedType == XingTableType.LatLong)
                        {
                            // Normal header-driven updater (program-generated tables)
                            UpdateLatLongTable(table, byKey, out matched, out updated, rowIndexMap, latLongContext);
                        }
                        else
                        {
                            // Try normal pass anyway
                            UpdateLatLongTable(table, byKey, out matched, out updated, rowIndexMap, latLongContext);

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
                .Where(r => !IsWaterCoordinateCrossingForPage(r))
                .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();

            // If this page only has water-coordinate crossings, do not create a blank PAGE table header.
            if (rows.Count == 0)
                return;

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
                        TrySetCellBlockScale(t, row, 0, CrossingBubbleScale);
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

        private static bool IsWaterCoordinateCrossingForPage(CrossingRecord record)
        {
            if (record == null)
                return false;

            var hasCoordinate = !string.IsNullOrWhiteSpace(record.Lat) ||
                                !string.IsNullOrWhiteSpace(record.Long);
            if (!hasCoordinate)
                return false;

            var owner = (record.Owner ?? string.Empty).Trim();
            owner = owner.Replace("\u2014", "-").Replace("\u2013", "-");

            return string.IsNullOrWhiteSpace(owner) || string.Equals(owner, "-", StringComparison.Ordinal);
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

            var regularStyleId = GetPreferredRegularTextStyle(db, tr, "Standard");
            var headerColor = Color.FromColorIndex(ColorMethod.ByAci, 254);
            var titleColor = Color.FromColorIndex(ColorMethod.ByAci, 14);

            if (showColumnHeaders)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    var cell = table.Cells[headerRow, c];
                    cell.TextHeight = CellTextHeight;
                    cell.Alignment = CellAlignment.MiddleCenter;
                    cell.TextStyleId = regularStyleId;
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
                titleCell.TextStyleId = regularStyleId;
                ApplyCellTextColor(titleCell, titleColor);
                ApplyTitleBorderStyle(table, titleRow, 0, table.Columns.Count - 1);
                ApplyTopBorderToRow(table, titleRow + 1, 0, table.Columns.Count - 1);
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
                    headerCell.TextStyleId = regularStyleId;
                    headerCell.TextString = headerText;
                    ApplyCellTextColor(headerCell, titleColor);
                    ApplyTitleBorderStyle(table, row, 0, table.Columns.Count - 1);
                    ApplyTopBorderToRow(table, row + 1, 0, table.Columns.Count - 1);
                    continue;
                }

                if (rowInfo.IsColumnHeader)
                {
                    ApplyLatLongColumnHeaderRow(table, row, HeaderRowHeight, CellTextHeight, regularStyleId, headerColor);
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
                SetBorderVisible(bordersObj, "Bottom", false);
                SetBorderVisible(bordersObj, "Outline", false);
            }
            catch
            {
                // purely cosmetic; ignore on unsupported releases
            }
        }

        private static void ApplyTopBorderToRow(Table table, int row, int startColumn, int endColumn)
        {
            if (table == null || table.Columns.Count == 0) return;
            if (row < 0 || row >= table.Rows.Count) return;

            startColumn = Math.Max(0, startColumn);
            endColumn = Math.Min(table.Columns.Count - 1, endColumn);

            for (int c = startColumn; c <= endColumn; c++)
            {
                ApplyTopBorder(table.Cells[row, c]);
            }
        }

        private static void ApplyTopBorder(Cell cell)
        {
            if (cell == null) return;

            try
            {
                var bordersProp = cell.GetType().GetProperty("Borders", BindingFlags.Public | BindingFlags.Instance);
                var bordersObj = bordersProp?.GetValue(cell, null);
                if (bordersObj == null) return;

                SetBorderVisible(bordersObj, "Top", true);
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

            // Header-free heuristics:
            // Keep this strict so unrelated tables with matching column counts are not misclassified.
            if (table.Columns.Count == 5)
            {
                if (HasMainHeaderRow(table) || HasExplicitCrossingKeys(table, minCount: 2))
                    return XingTableType.Main;
            }

            if (table.Columns.Count == 3)
            {
                if (HasPageHeaderRow(table) || HasExplicitCrossingKeys(table, minCount: 2))
                    return XingTableType.Page;
            }

            if ((table.Columns.Count == 4 || table.Columns.Count >= 6) && table.Rows.Count >= 1)
            {
                var headerColumns = Math.Min(table.Columns.Count, 6);
                if (HasHeaderRow(table, headerColumns, IsLatLongHeader) || LooksLikeLatLongTable(table))
                    return XingTableType.LatLong;
            }

            return XingTableType.Unknown;
        }

        private static bool HasExplicitCrossingKeys(Table table, int minCount)
        {
            if (table == null || table.Columns.Count <= 0 || minCount <= 0)
                return false;

            int found = 0;
            int maxRows = Math.Min(table.Rows.Count, 200);

            for (int row = 0; row < maxRows; row++)
            {
                var raw = ResolveCrossingKey(table, row, 0);
                if (string.IsNullOrWhiteSpace(raw))
                    continue;

                var normalized = NormalizeCrossingForMap(raw);
                if (Regex.IsMatch(normalized ?? string.Empty, @"^X0*\d+$"))
                {
                    found++;
                    if (found >= minCount)
                        return true;
                }
            }

            return false;
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

        private void UpdatePageTable(
            Table table,
            IDictionary<string, CrossingRecord> byKey,
            out int matched,
            out int updated,
            IDictionary<int, CrossingRecord> rowIndexMap = null,
            TableLayoutContext layoutContext = null)
        {
            matched = 0;
            updated = 0;
            if (table == null) return;

            if (!HasLayoutDwgRefContext(layoutContext))
            {
                Logger.Debug(_ed, $"table handle={table.ObjectId.Handle} type=PAGE status=skipped reason=no_layout_dwgref");
                return;
            }

            var columnCount = table.Columns.Count;
            var records = (byKey?.Values ?? Enumerable.Empty<CrossingRecord>())
                .Where(r => r != null)
                .ToList();

            // PAGE tables on XING#/CROSSING# layouts are DWG_REF-driven.
            var scopedRecords = FilterRecordsForPageTable(records, layoutContext);
            var topTr = table.Database?.TransactionManager?.TopTransaction as Transaction;

            // On crossing layouts, if a crossing is explicitly listed in a WATER LAT/LONG table
            // on the same layout, keep it only in that water table (exclude from PAGE table).
            var waterKeysOnLayout = GetWaterCrossingKeysOnLayout(table, topTr);
            if (waterKeysOnLayout.Count > 0)
            {
                var before = scopedRecords.Count;
                scopedRecords = scopedRecords
                    .Where(r =>
                    {
                        var key = NormalizeCrossingForMap(r?.Crossing);
                        return string.IsNullOrEmpty(key) || !waterKeysOnLayout.Contains(key);
                    })
                    .ToList();

                var removed = before - scopedRecords.Count;
                if (removed > 0)
                {
                    Logger.Debug(_ed, $"table handle={table.ObjectId.Handle} type=PAGE status=filtered_water duplicates_removed={removed}");
                }
            }

            var dataStartRow = GetDataStartRow(table, Math.Min(columnCount, 3), IsPageHeader);
            if (dataStartRow < 0) dataStartRow = 0;
            var bubbleBlockId = ResolvePageBubbleBlockId(table, topTr, dataStartRow);

            SyncTableDataRowsToCount(table, dataStartRow, scopedRecords.Count, columnCount, ref updated);

            for (var i = 0; i < scopedRecords.Count; i++)
            {
                var row = dataStartRow + i;
                if (row < 0 || row >= table.Rows.Count)
                    break;

                var record = scopedRecords[i];
                if (record == null)
                    continue;

                var rawKey = ResolveCrossingKey(table, row, 0);
                var logKey = !string.IsNullOrEmpty(rawKey) ? rawKey : (record.Crossing ?? string.Empty);
                matched++;

                var desiredCrossing = (record.Crossing ?? string.Empty).Trim();
                var desiredOwner = record.Owner ?? string.Empty;
                var desiredDescription = record.Description ?? string.Empty;

                var rowUpdated = false;

                if (columnCount > 0 && ValueDiffers(rawKey, desiredCrossing)) rowUpdated = true;
                if (columnCount > 1 && ValueDiffers(ReadCellText(table, row, 1), desiredOwner)) rowUpdated = true;
                if (columnCount > 2 && ValueDiffers(ReadCellText(table, row, 2), desiredDescription)) rowUpdated = true;

                // Always update the XING bubble even if OWNER/DESCRIPTION already match.
                if (columnCount > 0)
                {
                    var set = false;
                    if (!bubbleBlockId.IsNull && topTr != null)
                    {
                        try
                        {
                            set = TrySetCellToBlockWithAttribute(table, row, 0, topTr, bubbleBlockId, desiredCrossing);
                            if (set)
                            {
                                TrySetCellBlockScale(table, row, 0, CrossingBubbleScale);
                            }
                        }
                        catch { set = false; }
                    }

                    if (!set)
                    {
                        SetCellCrossingValue(table, row, 0, desiredCrossing);
                    }
                }

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

            // Best effort: if we couldn't delete surplus rows, clear any trailing content.
            for (var row = dataStartRow + scopedRecords.Count; row < table.Rows.Count; row++)
            {
                if (!RowHasAnyData(table, row, columnCount))
                    continue;

                if (TryClearRow(table, row, columnCount))
                {
                    updated++;
                    Logger.Debug(_ed, $"table handle={table.ObjectId.Handle} row={row} status=cleared reason=out_of_scope");
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
            IDictionary<int, CrossingRecord> rowIndexMap = null,
            TableLayoutContext layoutContext = null)
        {
            matched = 0;
            updated = 0;
            if (table == null) return;

            var columnCount = table.Columns.Count;
            var records = byKey?.Values ?? Enumerable.Empty<CrossingRecord>();
            var recordList = records.Where(r => r != null).ToList();

            // On XING#/CROSSING# layouts, LAT/LONG tables are enforced by DWG_REF.
            if (HasLayoutDwgRefContext(layoutContext))
            {
                SyncLatLongTableByDwgRef(table, recordList, layoutContext, out matched, out updated);
                RefreshTable(table);
                return;
            }

            var scopedRecords = FilterRecordsByLayoutContext(recordList, layoutContext);
            if (scopedRecords.Count == 0)
                scopedRecords = recordList;

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
                var rowLat = latColumn >= 0 ? ReadNorm(table, row, latColumn) : string.Empty;
                var rowLong = longColumn >= 0 ? ReadNorm(table, row, longColumn) : string.Empty;
                var hasCoordinates = !string.IsNullOrWhiteSpace(rowLat) && !string.IsNullOrWhiteSpace(rowLong);

                CrossingRecord record = null;

                // 1) Prefer stable row->record mapping from scan sources.
                if (rowIndexMap != null && rowIndexMap.TryGetValue(row, out var mapRecord))
                {
                    record = mapRecord;
                }

                // 2) Primary match by LAT/LONG (+ optional zone/dwg).
                if (record == null && hasCoordinates)
                {
                    record = FindRecordByLatLongColumns(table, row, scopedRecords, key);
                }

                // 3) Fallback to all records if context-scoped set had no coordinate match.
                if (record == null && hasCoordinates && !ReferenceEquals(scopedRecords, recordList))
                {
                    record = FindRecordByLatLongColumns(table, row, recordList, key);
                }

                // 4) Last resort: existing X# in the bubble.
                if (record == null && !string.IsNullOrEmpty(key) && byKey != null && byKey.TryGetValue(key, out var keyed))
                {
                    if (layoutContext == null || !layoutContext.HasFilters || RecordMatchesLayoutContext(keyed, layoutContext))
                    {
                        record = keyed;
                    }
                }

                if (record == null && !string.IsNullOrEmpty(key) && byKey != null)
                {
                    record = byKey.Values.FirstOrDefault(r =>
                        CrossingRecord.CompareCrossingKeys(r.Crossing, key) == 0 &&
                        (layoutContext == null || !layoutContext.HasFilters || RecordMatchesLayoutContext(r, layoutContext)));
                }

                if (record == null)
                {
                    Logger.Debug(_ed, $"Table {table.ObjectId.Handle}: no match for lat/long row {row} (x='{rawKey}', lat='{rowLat}', long='{rowLong}')");
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

        private void SyncLatLongTableByDwgRef(
            Table table,
            IList<CrossingRecord> records,
            TableLayoutContext layoutContext,
            out int matched,
            out int updated)
        {
            matched = 0;
            updated = 0;
            if (table == null)
                return;

            var columnCount = table.Columns.Count;
            if (columnCount < 4)
                return;

            var scopedRecords = FilterRecordsForLatLongTable(records, layoutContext);
            var dataStart = FindLatLongDataStartRow(table);
            if (dataStart < 0) dataStart = 0;

            SyncTableDataRowsToCount(table, dataStart, scopedRecords.Count, columnCount, ref updated);

            var hasExtendedLayout = columnCount >= 6;
            var zoneColumn = hasExtendedLayout ? 2 : -1;
            var latColumn = hasExtendedLayout ? 3 : 2;
            var longColumn = hasExtendedLayout ? 4 : 3;
            var dwgColumn = hasExtendedLayout ? 5 : -1;
            var remaining = new List<CrossingRecord>(scopedRecords);

            for (var i = 0; i < scopedRecords.Count; i++)
            {
                var row = dataStart + i;
                if (row < 0 || row >= table.Rows.Count)
                    break;

                var record = TakeNextLatLongRecordForRow(table, row, remaining, latColumn, longColumn);
                if (record == null)
                    continue;

                var rawKey = ResolveCrossingKey(table, row, 0);
                var desiredCrossing = (record.Crossing ?? string.Empty).Trim();
                var desiredDescription = record.Description ?? string.Empty;
                var desiredLat = record.Lat ?? string.Empty;
                var desiredLong = record.Long ?? string.Empty;
                var desiredZone = record.ZoneLabel ?? string.Empty;
                var desiredDwg = record.DwgRef ?? string.Empty;

                bool rowUpdated = false;

                if (columnCount > 0 && ValueDiffers(rawKey, desiredCrossing)) rowUpdated = true;
                if (columnCount > 1 && ValueDiffers(ReadCellText(table, row, 1), desiredDescription)) rowUpdated = true;
                if (zoneColumn >= 0 && ValueDiffers(ReadCellText(table, row, zoneColumn), desiredZone)) rowUpdated = true;
                if (latColumn >= 0 && ValueDiffers(ReadCellText(table, row, latColumn), desiredLat)) rowUpdated = true;
                if (longColumn >= 0 && ValueDiffers(ReadCellText(table, row, longColumn), desiredLong)) rowUpdated = true;
                if (dwgColumn >= 0 && ValueDiffers(ReadCellText(table, row, dwgColumn), desiredDwg)) rowUpdated = true;

                if (columnCount > 0) SetCellCrossingValue(table, row, 0, desiredCrossing);

                Cell descriptionCell = null;
                if (columnCount > 1) { try { descriptionCell = table.Cells[row, 1]; } catch { } SetCellValue(descriptionCell, desiredDescription); }

                if (zoneColumn >= 0)
                {
                    Cell zoneCell = null;
                    try { zoneCell = table.Cells[row, zoneColumn]; } catch { }
                    SetCellValue(zoneCell, desiredZone);
                }

                if (latColumn >= 0)
                {
                    Cell latCell = null;
                    try { latCell = table.Cells[row, latColumn]; } catch { }
                    SetCellValue(latCell, desiredLat);
                }

                if (longColumn >= 0)
                {
                    Cell longCell = null;
                    try { longCell = table.Cells[row, longColumn]; } catch { }
                    SetCellValue(longCell, desiredLong);
                }

                if (dwgColumn >= 0)
                {
                    Cell dwgCell = null;
                    try { dwgCell = table.Cells[row, dwgColumn]; } catch { }
                    SetCellValue(dwgCell, desiredDwg);
                }

                matched++;
                if (rowUpdated)
                {
                    updated++;
                    Logger.Debug(_ed, $"table handle={table.ObjectId.Handle} row={row} status=updated key={desiredCrossing}");
                }
            }

            // Best effort: if delete failed for surplus rows, clear trailing content.
            for (var row = dataStart + scopedRecords.Count; row < table.Rows.Count; row++)
            {
                if (!RowHasAnyData(table, row, columnCount))
                    continue;

                if (TryClearRow(table, row, columnCount))
                {
                    updated++;
                    Logger.Debug(_ed, $"table handle={table.ObjectId.Handle} row={row} status=cleared reason=out_of_scope");
                }
            }
        }

        private static void SyncTableDataRowsToCount(
            Table table,
            int dataStartRow,
            int desiredCount,
            int columnCount,
            ref int updatedCounter)
        {
            if (table == null)
                return;

            if (dataStartRow < 0) dataStartRow = 0;
            if (desiredCount < 0) desiredCount = 0;

            var keepFromRow = dataStartRow + desiredCount;
            for (var row = table.Rows.Count - 1; row >= keepFromRow; row--)
            {
                if (TryDeleteSingleRow(table, row))
                {
                    updatedCounter++;
                    continue;
                }

                if (TryClearRow(table, row, columnCount))
                {
                    updatedCounter++;
                }
            }

            var currentDataRows = Math.Max(0, table.Rows.Count - dataStartRow);
            var missing = desiredCount - currentDataRows;
            if (missing <= 0)
                return;

            var rowHeight = ResolveDataRowHeight(table, dataStartRow);
            if (TryInsertRows(table, table.Rows.Count, rowHeight, missing))
            {
                updatedCounter += missing;
            }
        }

        private static double ResolveDataRowHeight(Table table, int dataStartRow)
        {
            if (table == null)
                return 25.0;

            try
            {
                if (dataStartRow >= 0 && dataStartRow < table.Rows.Count)
                    return table.Rows[dataStartRow].Height;
            }
            catch { }

            try
            {
                if (table.Rows.Count > 0)
                    return table.Rows[table.Rows.Count - 1].Height;
            }
            catch { }

            return 25.0;
        }

        private static bool TryInsertRows(Table table, int rowIndex, double rowHeight, int count)
        {
            if (table == null || count <= 0)
                return false;

            try
            {
                table.InsertRows(rowIndex, rowHeight, count);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryDeleteSingleRow(Table table, int rowIndex)
        {
            if (table == null || rowIndex < 0 || rowIndex >= table.Rows.Count)
                return false;

            try
            {
                table.DeleteRows(rowIndex, 1);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryClearRow(Table table, int row, int columnCount)
        {
            if (table == null || row < 0 || row >= table.Rows.Count)
                return false;

            var changed = false;
            var cols = Math.Min(Math.Max(0, columnCount), table.Columns.Count);

            for (var col = 0; col < cols; col++)
            {
                var existing = ReadCellText(table, row, col);
                if (string.IsNullOrWhiteSpace(existing))
                    continue;

                if (col == 0)
                {
                    SetCellCrossingValue(table, row, col, string.Empty);
                }
                else
                {
                    Cell cell = null;
                    try { cell = table.Cells[row, col]; } catch { cell = null; }
                    SetCellValue(cell, string.Empty);
                }

                changed = true;
            }

            return changed;
        }

        private static bool RowHasAnyData(Table table, int row, int columnCount)
        {
            if (table == null || row < 0 || row >= table.Rows.Count)
                return false;

            var cols = Math.Min(Math.Max(0, columnCount), table.Columns.Count);
            for (var col = 0; col < cols; col++)
            {
                if (!string.IsNullOrWhiteSpace(ReadCellText(table, row, col)))
                    return true;
            }

            return false;
        }

        private ObjectId ResolvePageBubbleBlockId(Table table, Transaction tr, int dataStartRow)
        {
            if (table == null)
                return ObjectId.Null;

            var fromTable = TryGetBubbleBlockIdFromTable(table, dataStartRow);
            if (!fromTable.IsNull)
                return fromTable;

            if (tr == null || table.Database == null)
                return ObjectId.Null;

            try
            {
                return TryFindPrototypeBubbleBlockIdFromExistingTables(table.Database, tr);
            }
            catch
            {
                return ObjectId.Null;
            }
        }

        private static ObjectId TryGetBubbleBlockIdFromTable(Table table, int dataStartRow)
        {
            if (table == null || table.Columns.Count <= 0 || table.Rows.Count <= 0)
                return ObjectId.Null;

            var start = Math.Max(0, dataStartRow);
            var end = Math.Min(table.Rows.Count - 1, start + 20);

            for (var row = start; row <= end; row++)
            {
                Cell cell = null;
                try { cell = table.Cells[row, 0]; } catch { cell = null; }
                if (cell == null) continue;

                foreach (var content in EnumerateCellContentObjects(cell))
                {
                    var id = TryGetBlockTableRecordIdFromContent(content);
                    if (!id.IsNull)
                        return id;
                }
            }

            return ObjectId.Null;
        }

        private static CrossingRecord TakeNextLatLongRecordForRow(
            Table table,
            int row,
            IList<CrossingRecord> remaining,
            int latColumn,
            int longColumn)
        {
            if (remaining == null || remaining.Count == 0)
                return null;

            if (table != null && latColumn >= 0 && longColumn >= 0)
            {
                var rowLat = ReadNorm(table, row, latColumn);
                var rowLong = ReadNorm(table, row, longColumn);

                if (!string.IsNullOrWhiteSpace(rowLat) && !string.IsNullOrWhiteSpace(rowLong))
                {
                    for (var i = 0; i < remaining.Count; i++)
                    {
                        var candidate = remaining[i];
                        if (candidate == null) continue;

                        if (CoordinatesEquivalent(candidate.Lat, rowLat) &&
                            CoordinatesEquivalent(candidate.Long, rowLong))
                        {
                            remaining.RemoveAt(i);
                            return candidate;
                        }
                    }
                }
            }

            var first = remaining[0];
            remaining.RemoveAt(0);
            return first;
        }

        private HashSet<string> GetWaterCrossingKeysOnLayout(Table pageTable, Transaction tr)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (pageTable == null || tr == null)
                return keys;

            BlockTableRecord owner = null;
            try { owner = tr.GetObject(pageTable.OwnerId, OpenMode.ForRead, false) as BlockTableRecord; }
            catch { owner = null; }

            if (owner == null)
                return keys;

            foreach (ObjectId entId in owner)
            {
                if (entId == pageTable.ObjectId)
                    continue;

                if (!entId.ObjectClass.DxfName.Equals("ACAD_TABLE", StringComparison.OrdinalIgnoreCase))
                    continue;

                Table candidate = null;
                try { candidate = tr.GetObject(entId, OpenMode.ForRead, false) as Table; }
                catch { candidate = null; }
                if (candidate == null)
                    continue;

                XingTableType kind;
                try { kind = IdentifyTable(candidate, tr); }
                catch { kind = XingTableType.Unknown; }

                if (kind != XingTableType.LatLong)
                    continue;

                if (!IsWaterCrossingLatLongTable(candidate))
                    continue;

                foreach (var key in ExtractCrossingKeysFromLatLongTable(candidate))
                {
                    if (!string.IsNullOrWhiteSpace(key))
                        keys.Add(key);
                }
            }

            return keys;
        }

        private static bool IsWaterCrossingLatLongTable(Table table)
        {
            if (table == null || table.Rows.Count == 0 || table.Columns.Count == 0)
                return false;

            var maxRows = Math.Min(table.Rows.Count, 12);
            const string waterTitle = "WATERCROSSINGINFORMATION";

            for (var row = 0; row < maxRows; row++)
            {
                var text = ReadCellText(table, row, 0);
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                var normalized = NormalizeDwgRefForLookup(text);
                if (normalized.IndexOf(waterTitle, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }

            return false;
        }

        private static HashSet<string> ExtractCrossingKeysFromLatLongTable(Table table)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (table == null || table.Columns.Count <= 0 || table.Rows.Count <= 0)
                return keys;

            var columnCount = table.Columns.Count;
            var dataStart = FindLatLongDataStartRow(table);
            if (dataStart < 0) dataStart = 0;

            var hasExtendedLayout = columnCount >= 6;
            var zoneColumn = hasExtendedLayout ? 2 : -1;
            var latColumn = hasExtendedLayout ? 3 : 2;
            var longColumn = hasExtendedLayout ? 4 : 3;
            var dwgColumn = hasExtendedLayout ? 5 : -1;

            for (var row = dataStart; row < table.Rows.Count; row++)
            {
                if (IsLatLongSectionHeaderRow(table, row, zoneColumn, latColumn, longColumn, dwgColumn) ||
                    IsLatLongColumnHeaderRow(table, row))
                {
                    continue;
                }

                var raw = ResolveCrossingKey(table, row, 0);
                var normalized = NormalizeCrossingForMap(raw);
                if (string.IsNullOrWhiteSpace(normalized))
                    continue;

                if (!Regex.IsMatch(normalized, @"^X0*\d+$"))
                    continue;

                keys.Add(normalized);
            }

            return keys;
        }

        private static bool TableContainsAnyCrossingKeys(Table table, ISet<string> keys)
        {
            if (table == null || keys == null || keys.Count == 0)
                return false;

            foreach (var key in ExtractCrossingKeysFromLatLongTable(table))
            {
                if (!string.IsNullOrWhiteSpace(key) && keys.Contains(key))
                    return true;
            }

            return false;
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

        private static IDictionary<int, CrossingRecord> BuildPageRowIndexMap(
            ObjectId tableId,
            IEnumerable<CrossingRecord> records)
        {
            if (records == null)
                return new Dictionary<int, CrossingRecord>();

            return records
                .Where(r => r != null)
                .SelectMany(
                    r => (r.CrossingTableSources ?? new List<CrossingRecord.CrossingTableSource>())
                        .Where(s => s != null && !s.TableId.IsNull && s.TableId == tableId && s.RowIndex >= 0)
                        .Select(s => new { Source = s, Record = r }))
                .GroupBy(x => x.Source.RowIndex)
                .ToDictionary(g => g.Key, g => g.First().Record);
        }

        private static TableLayoutContext BuildTableLayoutContext(Table table)
        {
            var context = new TableLayoutContext();
            if (table == null || table.Database == null)
                return context;

            var tr = table.Database.TransactionManager?.TopTransaction as Transaction;
            if (tr == null)
                return context;

            BlockTableRecord owner = null;
            try { owner = tr.GetObject(table.OwnerId, OpenMode.ForRead, false) as BlockTableRecord; }
            catch { owner = null; }

            if (owner == null)
                return context;

            try
            {
                var layout = tr.GetObject(owner.LayoutId, OpenMode.ForRead, false) as Layout;
                context.LayoutName = layout?.LayoutName ?? string.Empty;
                context.DwgRef = TryExtractDwgRefFromLayoutName(context.LayoutName);
            }
            catch { }

            foreach (var text in EnumerateLayoutText(owner, tr))
            {
                foreach (var key in ExtractLocationKeys(text))
                {
                    if (!string.IsNullOrEmpty(key))
                        context.LocationKeys.Add(key);
                }
            }

            return context;
        }

        private static IEnumerable<string> EnumerateLayoutText(BlockTableRecord owner, Transaction tr)
        {
            if (owner == null || tr == null)
                yield break;

            foreach (ObjectId entId in owner)
            {
                Entity entity = null;
                try { entity = tr.GetObject(entId, OpenMode.ForRead, false) as Entity; }
                catch { entity = null; }

                if (entity == null)
                    continue;

                if (entity is DBText dbText)
                {
                    if (!string.IsNullOrWhiteSpace(dbText.TextString))
                        yield return dbText.TextString;
                    continue;
                }

                if (entity is MText mText)
                {
                    if (!string.IsNullOrWhiteSpace(mText.Text))
                        yield return mText.Text;
                    else if (!string.IsNullOrWhiteSpace(mText.Contents))
                        yield return mText.Contents;
                    continue;
                }

                if (entity is BlockReference br && br.AttributeCollection != null)
                {
                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        AttributeReference att = null;
                        try { att = tr.GetObject(attId, OpenMode.ForRead, false) as AttributeReference; }
                        catch { att = null; }

                        if (att == null)
                            continue;

                        if (att.IsMTextAttribute && att.MTextAttribute != null)
                        {
                            if (!string.IsNullOrWhiteSpace(att.MTextAttribute.Text))
                                yield return att.MTextAttribute.Text;
                            else if (!string.IsNullOrWhiteSpace(att.MTextAttribute.Contents))
                                yield return att.MTextAttribute.Contents;
                        }
                        else if (!string.IsNullOrWhiteSpace(att.TextString))
                        {
                            yield return att.TextString;
                        }
                    }
                }
            }
        }

        private static string TryExtractDwgRefFromLayoutName(string layoutName)
        {
            if (string.IsNullOrWhiteSpace(layoutName))
                return string.Empty;

            var match = LayoutDwgRefRegex.Match(layoutName);
            if (!match.Success)
                return string.Empty;

            return NormalizeDwgRefForLookup(match.Groups["ref"]?.Value);
        }

        private static string NormalizeDwgRefForLookup(string value)
        {
            var normalized = NormalizeKeyForLookup(value);
            return string.IsNullOrWhiteSpace(normalized) ? string.Empty : normalized;
        }

        private static List<CrossingRecord> FilterRecordsByLayoutContext(
            IEnumerable<CrossingRecord> records,
            TableLayoutContext context)
        {
            var list = (records ?? Enumerable.Empty<CrossingRecord>())
                .Where(r => r != null)
                .ToList();

            if (context == null || !context.HasFilters)
                return list;

            return list
                .Where(r => RecordMatchesLayoutContext(r, context))
                .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();
        }

        private static List<CrossingRecord> FilterRecordsForPageTable(
            IEnumerable<CrossingRecord> records,
            TableLayoutContext context)
        {
            return (records ?? Enumerable.Empty<CrossingRecord>())
                .Where(r => RecordMatchesLayoutDwgRef(r, context))
                .Where(r => !IsWaterCoordinateCrossingForPage(r))
                .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();
        }

        private static List<CrossingRecord> FilterRecordsForLatLongTable(
            IEnumerable<CrossingRecord> records,
            TableLayoutContext context)
        {
            return (records ?? Enumerable.Empty<CrossingRecord>())
                .Where(r => RecordMatchesLayoutDwgRef(r, context))
                .Where(r =>
                    !string.IsNullOrWhiteSpace(r?.Lat) &&
                    !string.IsNullOrWhiteSpace(r?.Long))
                .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();
        }

        private static bool RecordMatchesLayoutContext(CrossingRecord record, TableLayoutContext context)
        {
            if (record == null)
                return false;

            if (context == null || !context.HasFilters)
                return true;

            if (!string.IsNullOrWhiteSpace(context.DwgRef))
            {
                var recordDwg = NormalizeDwgRefForLookup(record.DwgRef);
                if (!string.Equals(recordDwg, context.DwgRef, StringComparison.OrdinalIgnoreCase))
                {
                    var suffixMatch = LayoutCloneSuffixRegex.Match(context.DwgRef);
                    var baseRef = suffixMatch.Success
                        ? NormalizeDwgRefForLookup(suffixMatch.Groups["base"]?.Value)
                        : string.Empty;

                    if (string.IsNullOrEmpty(baseRef) ||
                        !string.Equals(recordDwg, baseRef, StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }
                }
            }

            if (context.LocationKeys.Count > 0)
            {
                var locationKeys = BuildLocationKeyCandidates(record.Location);
                if (locationKeys.Count == 0)
                    return false;

                if (!locationKeys.Any(k => context.LocationKeys.Contains(k)))
                    return false;
            }

            return true;
        }

        private static bool HasLayoutDwgRefContext(TableLayoutContext context)
        {
            return context != null && !string.IsNullOrWhiteSpace(context.DwgRef);
        }

        private static bool RecordMatchesLayoutDwgRef(CrossingRecord record, TableLayoutContext context)
        {
            if (record == null)
                return false;

            // "-" / blank DWG_REF means this crossing is intentionally excluded from page tables.
            if (IsPlaceholderDwgRef(record.DwgRef))
                return false;

            if (!HasLayoutDwgRefContext(context))
                return false;

            var recordDwg = NormalizeDwgRefForLookup(record.DwgRef);
            if (string.IsNullOrWhiteSpace(recordDwg))
                return false;

            if (string.Equals(recordDwg, context.DwgRef, StringComparison.OrdinalIgnoreCase))
                return true;

            var suffixMatch = LayoutCloneSuffixRegex.Match(context.DwgRef);
            if (!suffixMatch.Success)
                return false;

            var baseRef = NormalizeDwgRefForLookup(suffixMatch.Groups["base"]?.Value);
            if (string.IsNullOrWhiteSpace(baseRef))
                return false;

            return string.Equals(recordDwg, baseRef, StringComparison.OrdinalIgnoreCase);
        }

        private static IEnumerable<string> ExtractLocationKeys(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                yield break;

            var cleaned = StripMTextFormatting(text);
            if (string.IsNullOrWhiteSpace(cleaned))
                yield break;

            foreach (Match match in MeridianLocationRegex.Matches(cleaned))
            {
                var key = NormalizeLocationKey(match.Value);
                if (!string.IsNullOrEmpty(key))
                    yield return key;
            }
        }

        private static List<string> BuildLocationKeyCandidates(string location)
        {
            var keys = new HashSet<string>(StringComparer.Ordinal);
            var direct = NormalizeLocationKey(location);
            if (!string.IsNullOrEmpty(direct))
                keys.Add(direct);

            if (!string.IsNullOrWhiteSpace(location))
            {
                try
                {
                    if (LocationFormatter.TryFormatMeridianLocation(location, out var formatted))
                    {
                        var formattedKey = NormalizeLocationKey(formatted);
                        if (!string.IsNullOrEmpty(formattedKey))
                            keys.Add(formattedKey);
                    }
                }
                catch { }
            }

            return keys.ToList();
        }

        private static string NormalizeLocationKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            var cleaned = StripMTextFormatting(value).Trim();
            if (string.IsNullOrEmpty(cleaned))
                return string.Empty;

            try
            {
                if (LocationFormatter.TryFormatMeridianLocation(cleaned, out var formatted) &&
                    !string.IsNullOrWhiteSpace(formatted))
                {
                    cleaned = formatted;
                }
            }
            catch { }

            var builder = new StringBuilder(cleaned.Length);
            foreach (var ch in cleaned)
            {
                if (char.IsLetterOrDigit(ch))
                    builder.Append(char.ToUpperInvariant(ch));
            }

            return builder.ToString();
        }

        private static bool IsPlaceholderDwgRef(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return true;

            var trimmed = value.Trim();
            trimmed = trimmed.Replace("\u2014", "-").Replace("\u2013", "-");
            trimmed = trimmed.Replace("Ã¢â‚¬â€", "-").Replace("Ã¢â‚¬â€œ", "-");
            trimmed = trimmed.Replace("ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Â", "-").Replace("ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Å“", "-");

            return string.Equals(trimmed, "-", StringComparison.Ordinal);
        }

        private static CrossingRecord TakeNextScopedRecord(
            IList<CrossingRecord> records,
            ISet<string> usedKeys,
            ref int cursor)
        {
            if (records == null)
                return null;

            while (cursor < records.Count)
            {
                var candidate = records[cursor++];
                if (candidate == null)
                    continue;

                var key = NormalizeCrossingForMap(candidate.Crossing);
                if (!string.IsNullOrEmpty(key) && usedKeys != null && usedKeys.Contains(key))
                    continue;

                return candidate;
            }

            return null;
        }

        private static void MarkScopedRecordUsed(CrossingRecord record, ISet<string> usedKeys)
        {
            if (record == null || usedKeys == null)
                return;

            var key = NormalizeCrossingForMap(record.Crossing);
            if (!string.IsNullOrEmpty(key))
                usedKeys.Add(key);
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

        private static ObjectId GetPreferredRegularTextStyle(Database db, Transaction tr, string preferredStyleName)
        {
            if (db == null || tr == null)
                return ObjectId.Null;

            var tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
            if (!string.IsNullOrWhiteSpace(preferredStyleName) && tst.Has(preferredStyleName))
                return tst[preferredStyleName];

            if (!db.Textstyle.IsNull)
                return db.Textstyle;

            foreach (ObjectId id in tst)
                return id;

            return ObjectId.Null;
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
        // Ensures autoscale is OFF and final scale == CrossingBubbleScale.
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
                // Force configured bubble scale after placement (table-level API if available, else per-content)
                TrySetCellBlockScale(table, row, col, CrossingBubbleScale);
                TrySetCellMiddleCenter(cell);
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

            Cell cell = null;
            try { cell = t.Cells[row, col]; } catch { cell = null; }
            TrySetCellMiddleCenter(cell);

            // 1) Try to set the block attribute by any of our known tags
            foreach (var tag in CrossingAttributeTags)
            {
                if (TrySetBlockAttributeValue(t, row, col, tag, crossingText))
                    return;
            }

            // 2) If the cell currently hosts a block, do NOT overwrite it with text.
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

        internal static bool TrySetCellCrossingValue(Table table, int row, int col, string crossingText)
        {
            if (table == null)
                return false;

            var before = NormalizeCrossingForMap(ResolveCrossingKey(table, row, col));
            SetCellCrossingValue(table, row, col, crossingText);
            var after = NormalizeCrossingForMap(ResolveCrossingKey(table, row, col));
            var desired = NormalizeCrossingForMap(crossingText);

            if (!string.IsNullOrEmpty(desired) && string.Equals(after, desired, StringComparison.OrdinalIgnoreCase))
                return true;

            if (string.IsNullOrEmpty(before) && string.IsNullOrEmpty(after))
                return false;

            return !string.Equals(before, after, StringComparison.OrdinalIgnoreCase);
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
            TrySetCellMiddleCenter(cell);

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

        private static void TrySetCellMiddleCenter(Cell cell)
        {
            if (cell == null)
                return;

            try
            {
                cell.Alignment = CellAlignment.MiddleCenter;
            }
            catch
            {
                // best effort: some styles/cells may reject alignment writes
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
        // Try to set the scale of a block hosted in a table cell to a specific value (e.g., 0.980)
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

        private static CrossingRecord FindRecordByLatLongColumns(
            Table t,
            int row,
            IEnumerable<CrossingRecord> records,
            string normalizedKey = null)
        {
            if (t == null || records == null)
                return null;

            var columns = t?.Columns.Count ?? 0;
            var zone = columns >= 6 ? ReadNorm(t, row, 2) : string.Empty;
            var lat = ReadNorm(t, row, columns >= 6 ? 3 : 2);
            var lng = ReadNorm(t, row, columns >= 6 ? 4 : 3);
            var dwg = columns >= 6 ? ReadNorm(t, row, 5) : string.Empty;

            if (string.IsNullOrWhiteSpace(lat) || string.IsNullOrWhiteSpace(lng))
                return null;

            var candidates = records.Where(r =>
                CoordinatesEquivalent(r?.Lat, lat) &&
                CoordinatesEquivalent(r?.Long, lng) &&
                (string.IsNullOrEmpty(zone) || string.Equals(Norm(r?.Zone), zone, StringComparison.OrdinalIgnoreCase)) &&
                (string.IsNullOrEmpty(dwg) || string.Equals(Norm(r?.DwgRef), dwg, StringComparison.OrdinalIgnoreCase))).ToList();
            if (candidates.Count == 1) return candidates[0];

            if (candidates.Count > 1 && !string.IsNullOrEmpty(normalizedKey))
            {
                var keyed = candidates.FirstOrDefault(r =>
                    string.Equals(NormalizeKeyForLookup(r?.Crossing), normalizedKey, StringComparison.OrdinalIgnoreCase) ||
                    CrossingRecord.CompareCrossingKeys(r?.Crossing, normalizedKey) == 0);
                if (keyed != null)
                    return keyed;
            }

            return null;
        }

        private static bool CoordinatesEquivalent(string left, string right)
        {
            var l = Norm(left);
            var r = Norm(right);
            if (string.IsNullOrEmpty(l) || string.IsNullOrEmpty(r))
                return string.Equals(l, r, StringComparison.OrdinalIgnoreCase);

            if (double.TryParse(l, NumberStyles.Float, CultureInfo.InvariantCulture, out var ld) &&
                double.TryParse(r, NumberStyles.Float, CultureInfo.InvariantCulture, out var rd))
            {
                return Math.Abs(ld - rd) <= 0.000001d;
            }

            return string.Equals(l, r, StringComparison.OrdinalIgnoreCase);
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
