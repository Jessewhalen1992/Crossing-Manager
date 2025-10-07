using Autodesk.AutoCAD.ApplicationServices;  // Document, CommandEventArgs
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using XingManager.Models;
using XingManager.Services;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using WinFormsFlowDirection = System.Windows.Forms.FlowDirection;

namespace XingManager
{
    public partial class XingForm : UserControl
    {
        private readonly Document _doc;
        private readonly XingRepository _repository;
        private readonly TableSync _tableSync;
        private readonly LayoutUtils _layoutUtils;
        private readonly TableFactory _tableFactory;
        private readonly Serde _serde;
        private readonly DuplicateResolver _duplicateResolver;
        private readonly LatLongDuplicateResolver _latLongDuplicateResolver;

        private BindingList<CrossingRecord> _records = new BindingList<CrossingRecord>();
        private IDictionary<ObjectId, DuplicateResolver.InstanceContext> _contexts =
            new Dictionary<ObjectId, DuplicateResolver.InstanceContext>();

        private bool _isDirty;
        private bool _isScanning;
        private bool _isAwaitingRenumber;
        private int? _currentUtmZone;
        private bool _isUpdatingZoneControl;

        private const string TemplatePath = @"M:\Drafting\_CURRENT TEMPLATES\Compass_Main.dwt";
        private const string DefaultTemplateLayoutName = "X";
        private const string HydroTemplateLayoutName = "H2O-PROFILE";
        private const string CnrTemplateLayoutName = "CNR-PROFILE";
        private const string HighwayTemplateLayoutName = "HWY-PROFILE";
        private static readonly string[] HydroKeywords = { "Watercourse", "Creek", "River" };
        private static readonly string[] HydroOwnerKeywords =
        {
            "Nova",
            "Alliance",
            "Pembina",
            "TCPL",
            "Ovintiv",
            "PGI"
        };
        private static readonly string[] RailwayKeywords = { "Railway" };
        private static readonly string[] HighwayKeywords = { "Highway", "Hwy" };
        private const string CreateAllPagesDisplayText = "Create ALL XING pages...";
        private static readonly IComparer<string> DwgRefComparer = new NaturalDwgRefComparer();

        public XingForm(
            Document doc,
            XingRepository repository,
            TableSync tableSync,
            LayoutUtils layoutUtils,
            TableFactory tableFactory,
            Serde serde,
            DuplicateResolver duplicateResolver,
            LatLongDuplicateResolver latLongDuplicateResolver)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (repository == null) throw new ArgumentNullException(nameof(repository));
            if (tableSync == null) throw new ArgumentNullException(nameof(tableSync));
            if (layoutUtils == null) throw new ArgumentNullException(nameof(layoutUtils));
            if (tableFactory == null) throw new ArgumentNullException(nameof(tableFactory));
            if (serde == null) throw new ArgumentNullException(nameof(serde));
            if (duplicateResolver == null) throw new ArgumentNullException(nameof(duplicateResolver));
            if (latLongDuplicateResolver == null) throw new ArgumentNullException(nameof(latLongDuplicateResolver));

            InitializeComponent();

            _doc = doc;
            _repository = repository;
            _tableSync = tableSync;
            _layoutUtils = layoutUtils;
            _tableFactory = tableFactory;
            _serde = serde;
            _duplicateResolver = duplicateResolver;
            _latLongDuplicateResolver = latLongDuplicateResolver;

            ConfigureGrid();
            UpdateZoneControlFromState();
        }

        // ===== Public entry points called by your commands =====

        public void LoadData() => RescanRecords();

        public void RescanData() => RescanRecords();

        public void ApplyToDrawing() => ApplyChangesToDrawing();

        public void GenerateXingPageFromCommand() => GenerateXingPage();

        public void GenerateAllXingPagesFromCommand() => GenerateAllXingPages();

        public void GenerateAllLatLongTablesFromCommand() => GenerateAllLatLongTables();

        public void CreateLatLongRowFromCommand() => CreateOrUpdateLatLongTable();

        private void OnCommandMatchTableDone(object sender, CommandEventArgs e)
        {
            try
            {
                if (!string.Equals(e.GlobalCommandName, "XING_MATCH_TABLE", StringComparison.OrdinalIgnoreCase))
                    return;

                if (sender is Document d)
                {
                    d.CommandEnded -= OnCommandMatchTableDone;
                    d.CommandCancelled -= OnCommandMatchTableDone;
                    try { d.CommandFailed -= OnCommandMatchTableDone; } catch { }
                }

                // Refresh grid only (no table writes)
                RescanRecords(applyToTables: false);
            }
            catch { /* best effort */ }
        }

        public void RenumberSequentiallyFromCommand()
        {
            StartRenumberCrossingCommand();
        }

        public void AddRncPolylineFromCommand() => AddRncPolyline();

        // ===== UI wiring =====

        private void ConfigureGrid()
        {
            gridCrossings.AutoGenerateColumns = false;
            gridCrossings.Columns.Clear();

            gridCrossings.Columns.Add(CreateTextColumn("Crossing", "CROSSING", 80));
            gridCrossings.Columns.Add(CreateTextColumn("Owner", "OWNER", 120));
            gridCrossings.Columns.Add(CreateTextColumn("Description", "DESCRIPTION", 200));
            gridCrossings.Columns.Add(CreateTextColumn("Location", "LOCATION", 200));
            gridCrossings.Columns.Add(CreateTextColumn("DwgRef", "DWG_REF", 100));
            gridCrossings.Columns.Add(CreateTextColumn("Lat", "LAT", 100));
            gridCrossings.Columns.Add(CreateTextColumn("Long", "LONG", 100));

            gridCrossings.DataSource = _records;
            gridCrossings.CellValueChanged += GridCrossingsOnCellValueChanged;
            gridCrossings.CurrentCellDirtyStateChanged += GridCrossingsOnCurrentCellDirtyStateChanged;
            gridCrossings.CellFormatting += GridCrossingsOnCellFormatting;
        }

        private static DataGridViewTextBoxColumn CreateTextColumn(string propertyName, string header, int width)
        {
            return new DataGridViewTextBoxColumn
            {
                DataPropertyName = propertyName,
                HeaderText = header,
                Width = width,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
        }

        private void GridCrossingsOnCurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (gridCrossings.IsCurrentCellDirty)
                gridCrossings.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void GridCrossingsOnCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_isScanning) return;
            _isDirty = true;
        }

        private void GridCrossingsOnCellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e == null || e.ColumnIndex < 0)
                return;

            var column = gridCrossings.Columns[e.ColumnIndex];
            if (column == null)
                return;

            if (!string.Equals(column.DataPropertyName, nameof(CrossingRecord.Zone), StringComparison.Ordinal))
                return;

            if (e.Value is string zone && !string.IsNullOrWhiteSpace(zone))
            {
                e.Value = string.Format(CultureInfo.InvariantCulture, "ZONE {0}", zone.Trim());
                e.FormattingApplied = true;
            }
        }

        // ===== Buttons (Designer wires to these; keep them) =====

        private void btnRescan_Click(object sender, EventArgs e) => RescanRecords();

        private void btnApply_Click(object sender, EventArgs e) => ApplyChangesToDrawing();

        private void btnGenerateAllPages_Click(object sender, EventArgs e) => GenerateAllXingPages();

        private void btnGenerateAllLatLongTables_Click(object sender, EventArgs e) => GenerateAllLatLongTables();

        // DELETE SELECTED: does NOT renumber the remaining crossings.
        // - removes the block instances for the selected record
        // - removes matching rows from all recognized tables (Main/Page/LatLong)
        // - writes current grid back to blocks
        // - forces a real graphics refresh to avoid the “ghost row”
        private void btnDelete_Click(object sender, EventArgs e)
        {
            var record = GetSelectedRecord();
            if (record == null)
            {
                MessageBox.Show("Select a crossing to delete.", "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var confirm = MessageBox.Show(
                $"Delete crossing {record.Crossing} (blocks + any table rows)?",
                "Crossing Manager", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes) return;

            var crossingKey = record.Crossing; // keep before removal

            try
            {
                // 1) Delete the block instances for this record
                _repository.DeleteInstances(record.AllInstances);

                // 2) Remove from the grid (NO renumbering)
                _records.Remove(record);

                // 3) Persist current grid values back to blocks
                _repository.ApplyChanges(_records.ToList(), _tableSync);

                // 4) Remove row(s) from every recognized crossing table
                DeleteRowFromTables(crossingKey);

                // 5) Update remaining table cells from grid (B..E only; never touches col 0)
                UpdateAllXingTablesFromGrid();

                // 6) Extra safety: force redraw/regen of all crossing tables
                ForceRegenAllCrossingTablesInDwg();

                _isDirty = true;
                gridCrossings.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnRenumber_Click(object sender, EventArgs e)
        {
            StartRenumberCrossingCommand();
        }

        private void btnAddRncPolyline_Click(object sender, EventArgs e)
        {
            AddRncPolyline();
        }

        // Button: MATCH TABLE  ->  Merge from selected table, then persist to blocks.
        // No table writes here.
        private void btnMatchTable_Click(object sender, EventArgs e)
        {
            // Read a selected table into the grid only
            var gridChanged = MatchTableIntoGrid();   // <- now returns bool

            if (gridChanged)
            {
                try
                {
                    // Push grid -> crossing block attributes
                    _repository.ApplyChanges(_records.ToList(), _tableSync);

                    // OPTIONAL: also push grid -> tables right away (uncomment if you want that here too)
                    // UpdateAllXingTablesFromGrid();

                    // OPTIONAL: refresh grid from DWG (no table writes)
                    // RescanRecords(applyToTables: false);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void MatchTableFromCommand()
        {
            // Same behavior when invoked as a command
            MatchTableIntoGrid(persistAfterMatch: true);
        }
        private void btnGeneratePage_Click(object sender, EventArgs e) => GenerateXingPage();

        private void btnLatLong_Click(object sender, EventArgs e) => CreateOrUpdateLatLongTable();

        private void btnAddLatLong_Click(object sender, EventArgs e) => AddLatLongFromDrawing();

        private void btnExport_Click(object sender, EventArgs e)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                dialog.FileName = "Crossings.csv";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _serde.Export(dialog.FileName, _records);
                        MessageBox.Show("Export complete.", "Crossing Manager",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Crossing Manager",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var imported = _serde.Import(dialog.FileName);
                        MergeImportedRecords(imported);
                        _isDirty = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Crossing Manager",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // ===== Core actions =====

        private void RescanRecords(bool applyToTables = true, bool suppressDuplicateUi = false)
        {
            _isScanning = true;
            try
            {
                // NEW: in case Rescan is called while a cell is still being edited
                FlushPendingGridEdits();

                // 1) Read from DWG -> populate grid
                var firstScan = _repository.ScanCrossings();
                _records = new BindingList<CrossingRecord>(firstScan.Records.ToList());
                _contexts = firstScan.InstanceContexts ?? new Dictionary<ObjectId, DuplicateResolver.InstanceContext>();
                gridCrossings.DataSource = _records;

                // 2) Resolve duplicates unless explicitly suppressed
                bool dupOk = true, latDupOk = true;
                if (!suppressDuplicateUi)
                {
                    dupOk = _duplicateResolver.ResolveDuplicates(_records, _contexts);          // :contentReference[oaicite:4]{index=4}
                    latDupOk = _latLongDuplicateResolver.ResolveDuplicates(_records, _contexts);
                    if (!dupOk || !latDupOk)
                    {
                        _records?.ResetBindings();
                        gridCrossings.Refresh();
                        _isDirty = false;
                        return;
                    }
                }

                _records.ResetBindings();
                gridCrossings.Refresh();

                // 3) Optionally persist to DWG/tables (used by a true "Rescan" merge pass)
                if (applyToTables)
                {
                    var snapshot = _records.ToList();

                    try
                    {
                        _repository.ApplyChanges(snapshot, _tableSync);
                        _tableSync.UpdateAllTables(_doc, snapshot);  // MAIN/PAGE/LATLONG  :contentReference[oaicite:5]{index=5}
                        UpdateAllXingTablesFromGrid();               // safety X‑key pass (4/6‑col LAT/LONG)
                        _tableSync.UpdateLatLongSourceTables(_doc, snapshot); // explicit sources  :contentReference[oaicite:6]{index=6}
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Crossing Manager",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    // 4) Re-scan to reflect the persisted state (no resolver here)
                    var post = _repository.ScanCrossings();
                    _records = new BindingList<CrossingRecord>(post.Records.ToList());
                    _contexts = post.InstanceContexts ?? new Dictionary<ObjectId, DuplicateResolver.InstanceContext>();
                    gridCrossings.DataSource = _records;
                    _records.ResetBindings();
                    gridCrossings.Refresh();
                }

                _isDirty = false;
                UpdateZoneSelectionFromRecords();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _isScanning = false;
            }
        }

        private void ApplyChangesToDrawing()
        {
            // NEW: make sure the last keystroke in the grid is committed
            FlushPendingGridEdits();

            if (!ValidateRecords()) return;

            try
            {
                var snapshot = _records.ToList();

                // 1) Persist grid -> blocks
                _repository.ApplyChanges(snapshot, _tableSync);

                // 2) Persist grid -> ALL recognised tables (MAIN / PAGE / LATLONG)
                _tableSync.UpdateAllTables(_doc, snapshot);   // (uses robust header/heuristics)  :contentReference[oaicite:2]{index=2}

                // 3) Safety pass: force an X‑key keyed update for every recognised table
                //    (handles both 4‑col and 6‑col LAT/LONG layouts).  See method below.
                UpdateAllXingTablesFromGrid();

                // 4) If we captured explicit LAT/LONG sources on scan, update those rows by id as well
                _tableSync.UpdateLatLongSourceTables(_doc, snapshot);  // :contentReference[oaicite:3]{index=3}

                // 5) Rebuild table display lists so the change is visible immediately
                ForceRegenAllCrossingTablesInDwg();

                // 6) Re-read **without** invoking the duplicate resolver UI
                RescanRecords(applyToTables: false, suppressDuplicateUi: true);

                _isDirty = false;
                MessageBox.Show("Crossing data applied to drawing.", "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool ValidateRecords()
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var record in _records)
            {
                if (string.IsNullOrWhiteSpace(record.Crossing))
                {
                    MessageBox.Show("Each record must have a CROSSING value.", "Crossing Manager",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                var key = record.Crossing.Trim().ToUpperInvariant();
                if (!seen.Add(key))
                {
                    MessageBox.Show(
                        string.Format(CultureInfo.InvariantCulture, "Duplicate CROSSING value '{0}' detected.", record.Crossing),
                        "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (!ValidateLatLongValue(record.Lat) || !ValidateLatLongValue(record.Long))
                {
                    MessageBox.Show(
                        string.Format(CultureInfo.InvariantCulture, "LAT/LONG values for {0} must be decimal numbers.", record.Crossing),
                        "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            return true;
        }

        private static bool ValidateLatLongValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return true;
            double parsed;
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed);
        }

        private void MergeImportedRecords(IEnumerable<CrossingRecord> imported)
        {
            var map = _records.ToDictionary(r => r.CrossingKey, r => r, StringComparer.OrdinalIgnoreCase);
            foreach (var record in imported)
            {
                if (map.TryGetValue(record.CrossingKey, out var existing))
                {
                    existing.Owner = record.Owner;
                    existing.Description = record.Description;
                    existing.Location = record.Location;
                    existing.DwgRef = record.DwgRef;
                    existing.Lat = record.Lat;
                    existing.Long = record.Long;
                    if (!string.IsNullOrWhiteSpace(record.Zone))
                    {
                        existing.Zone = record.Zone;
                    }
                }
                else
                {
                    _records.Add(record);
                }
            }

            UpdateZoneSelectionFromRecords();
        }

        // ===== Update all recognized crossing tables from the grid =====

        /// Update every recognized crossing table (MAIN, PAGE, LATLONG) from the current grid.
        /// Matching uses the same X-key logic as MatchTableIntoGrid: attribute-first with fallbacks,
        /// and NormalizeXKey handles digits-only (e.g., "3" => "X3").
        /// Update recognized crossing tables (MAIN/PAGE/LATLONG) from the current grid,
        /// but never modify Column A (X/bubble). Columns B..E only.
        /// <summary>
        /// Update recognised crossing tables (MAIN, PAGE, LATLONG) from the current grid.
        /// Never modifies column A (bubble index); updates only columns B..E accordingly.
        /// </summary>
        /// Update recognized crossing tables (MAIN/PAGE/LATLONG) from the current grid.
        /// Never modifies Column 0 (the bubble/X). Only columns B..E are written.
        /// Ends with a hard refresh to avoid the transient “ghost row” look.
        /// Update recognised crossing tables (MAIN/PAGE/LATLONG) from the current grid.
        /// Never modifies Column 0 (the bubble/X); only columns B..E are written.
        /// Ends with a hard refresh to avoid the transient “ghost row” look.
        /// Update recognised crossing tables (MAIN, PAGE, LATLONG) from the current grid.
        /// - Matches rows by X-key (attribute-first, with text fallback).
        /// - Never writes Column 0 (bubble) in any table.
        /// - For LAT/LONG, supports both 4‑col (ID, DESC, LAT, LONG) and 6‑col (ID, DESC, ZONE, LAT, LONG, DWG_REF).
        /// - Ends with a regen so changes appear immediately.
        private void UpdateAllXingTablesFromGrid()
        {
            var doc = _doc ?? AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            // Index current records by normalized X key ("3" -> "X3")
            var byX = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in _records)
            {
                var key = NormalizeXKey(r.Crossing);
                if (!string.IsNullOrWhiteSpace(key) && !byX.ContainsKey(key))
                    byX[key] = r;
            }

            using (doc.LockDocument())
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    foreach (ObjectId entId in btr)
                    {
                        if (!entId.ObjectClass.DxfName.Equals("ACAD_TABLE", StringComparison.OrdinalIgnoreCase))
                            continue;

                        var table = tr.GetObject(entId, OpenMode.ForRead) as Table;
                        if (table == null) continue;

                        var kind = _tableSync.IdentifyTable(table, tr); // robust classifier  :contentReference[oaicite:7]{index=7}
                        if (kind == TableSync.XingTableType.Unknown) continue;

                        table.UpgradeOpen();

                        int matched = 0, updated = 0;

                        // Data start row (LAT/LONG tables may have a title + header)
                        int startRow = 0;
                        if (kind == TableSync.XingTableType.LatLong)
                        {
                            try { startRow = TableSync.FindLatLongDataStartRow(table); } // header-aware  :contentReference[oaicite:8]{index=8}
                            catch { startRow = 0; }
                            if (startRow < 0) startRow = 0;
                        }

                        for (int row = startRow; row < table.Rows.Count; row++)
                        {
                            // Column 0 key: attribute-first, fall back to visible token
                            string xRaw = ReadXFromCellAttributeOnly(table, row, tr);
                            if (string.IsNullOrWhiteSpace(xRaw))
                            {
                                try
                                {
                                    var txt = table.Cells[row, 0]?.TextString ?? string.Empty;
                                    xRaw = ExtractXToken(txt);
                                }
                                catch { }
                            }

                            var xKey = NormalizeXKey(xRaw);
                            if (string.IsNullOrWhiteSpace(xKey) || !byX.TryGetValue(xKey, out var rec))
                            {
                                ed?.WriteMessage($"\n[CrossingManager] Row {row} -> NO MATCH (key='{xRaw}')");
                                continue;
                            }

                            matched++;
                            bool changed = false;

                            // Never touch Column 0 (bubble)
                            if (kind == TableSync.XingTableType.Main)
                            {
                                changed |= SetCellIfChanged(table, row, 1, rec.Owner);
                                changed |= SetCellIfChanged(table, row, 2, rec.Description);
                                changed |= SetCellIfChanged(table, row, 3, rec.Location);
                                changed |= SetCellIfChanged(table, row, 4, rec.DwgRef);
                            }
                            else if (kind == TableSync.XingTableType.Page)
                            {
                                changed |= SetCellIfChanged(table, row, 1, rec.Owner);
                                changed |= SetCellIfChanged(table, row, 2, rec.Description);
                            }
                            else if (kind == TableSync.XingTableType.LatLong)
                            {
                                // 4‑col: ID | DESC | LAT | LONG
                                // 6‑col: ID | DESC | ZONE | LAT | LONG | DWG_REF
                                var cols = table.Columns.Count;

                                // Description (col 1 is common)
                                changed |= SetCellIfChanged(table, row, 1, rec.Description);

                                // Zone label for 6‑col tables (e.g., "ZONE 11")
                                var zoneLabel = string.IsNullOrWhiteSpace(rec.Zone)
                                    ? string.Empty
                                    : string.Format(CultureInfo.InvariantCulture, "ZONE {0}", rec.Zone.Trim());

                                if (cols >= 6)
                                {
                                    changed |= SetCellIfChanged(table, row, 2, zoneLabel);
                                    changed |= SetCellIfChanged(table, row, 3, rec.Lat);
                                    changed |= SetCellIfChanged(table, row, 4, rec.Long);
                                    changed |= SetCellIfChanged(table, row, 5, rec.DwgRef);
                                }
                                else
                                {
                                    changed |= SetCellIfChanged(table, row, 2, rec.Lat);
                                    changed |= SetCellIfChanged(table, row, 3, rec.Long);
                                }
                            }

                            if (changed) updated++;
                        }

                        if (updated > 0)
                        {
                            try { table.GenerateLayout(); } catch { }
                            ForceRegenTable(table);
                        }

                        var handleHex = table.ObjectId.Handle.Value.ToString("X");
                        ed?.WriteMessage($"\n[CrossingManager] Table {handleHex}: {kind.ToString().ToUpperInvariant()} matched={matched} updated={updated}");
                    }
                }

                tr.Commit();
            }

            try { doc.Editor?.Regen(); } catch { }
            try { AcadApp.UpdateScreen(); } catch { }
        }

        /// Set a table cell's TextString only if it actually changes; returns true if changed.
        private static bool SetCellIfChanged(Table t, int row, int col, string newText)
        {
            if (t == null) return false;

            var desired = newText ?? string.Empty;

            Cell cell = null;
            try { cell = t.Cells[row, col]; }
            catch { cell = null; }

            if (cell == null) return false;

            string current;
            try { current = cell.TextString ?? string.Empty; }
            catch { current = string.Empty; }

            if (string.Equals(current, desired, StringComparison.Ordinal))
                return false;

            try
            {
                cell.TextString = desired;
                return true;
            }
            catch
            {
                try
                {
                    cell.Value = desired;
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }

        // ===== MATCH TABLE (grid-only) =====

        /// MATCH TABLE:
        /// - Reads a selected table and merges values into the grid (never writes DWG/table here)
        /// - Returns true if the grid changed.
        /// - If persistAfterMatch = true, also writes grid -> blocks and grid -> tables afterwards.
        private bool MatchTableIntoGrid(bool persistAfterMatch = false)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return false;

            var ed = doc.Editor;
            var peo = new PromptEntityOptions("\nSelect a crossing table (Main/Page):");
            peo.SetRejectMessage("\nEntity must be a TABLE.");
            peo.AddAllowedClass(typeof(Table), exactMatch: true);

            var sel = ed.GetEntity(peo);
            if (sel.Status != PromptStatus.OK) return false;

            bool gridChanged = false;

            // === Read the table and update the grid (grid only) ===
            using (doc.LockDocument())
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var table = tr.GetObject(sel.ObjectId, OpenMode.ForRead) as Table;
                if (table == null)
                {
                    ed.WriteMessage("\n[CrossingManager] Selection was not a Table.");
                    return false;
                }

                var kind = _tableSync.IdentifyTable(table, tr);
                if (kind == TableSync.XingTableType.Unknown)
                {
                    ed.WriteMessage("\n[CrossingManager] Could not determine table type.");
                    return false;
                }
                if (kind == TableSync.XingTableType.LatLong)
                {
                    ed.WriteMessage("\n[CrossingManager] Lat/Long tables are not supported by MATCH TABLE.");
                    return false;
                }

                var byX = new Dictionary<string, (string Owner, string Desc, string Loc, string Dwg, int Row)>(StringComparer.OrdinalIgnoreCase);

                ed.WriteMessage("\n[CrossingManager] --- TABLE VALUES (parsed) ---");
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    // Column A: attribute-first X-key (strict)
                    string xRaw = ReadXFromCellAttributeOnly(table, r, tr);
                    if (string.IsNullOrWhiteSpace(xRaw)) continue;

                    var owner = ReadTableCellText(table, r, 1);
                    var desc = ReadTableCellText(table, r, 2);

                    string loc = string.Empty, dwg = string.Empty;
                    if (kind == TableSync.XingTableType.Main)
                    {
                        loc = ReadTableCellText(table, r, 3);
                        dwg = ReadTableCellText(table, r, 4);
                    }

                    var xKey = NormalizeXKey(xRaw);
                    ed.WriteMessage($"\n[CrossingManager] [T] row {r}: X='{xKey}' owner='{owner}' desc='{desc}' loc='{loc}' dwg='{dwg}'");

                    if (!byX.ContainsKey(xKey))
                        byX[xKey] = (owner, desc, loc, dwg, r);
                    else
                        ed.WriteMessage($"\n[CrossingManager] [!] Duplicate X '{xKey}' in table (row {r}). First occurrence kept.");
                }

                ed.WriteMessage($"\n[CrossingManager] Table rows indexed by X = {byX.Count}");
                ed.WriteMessage("\n[CrossingManager] --- TABLE→FORM UPDATES (X-only) ---");

                int matched = 0, updated = 0, noMatch = 0;

                foreach (var rec in _records)
                {
                    var xKey = NormalizeXKey(rec.Crossing);
                    if (string.IsNullOrWhiteSpace(xKey) || !byX.TryGetValue(xKey, out var src))
                    {
                        ed.WriteMessage($"\n[CrossingManager] [!] {rec.Crossing}: no matching X in table.");
                        noMatch++;
                        continue;
                    }

                    bool changed = false;

                    if (!string.Equals(rec.Owner, src.Owner, StringComparison.Ordinal))
                    { rec.Owner = src.Owner; changed = true; }

                    if (!string.Equals(rec.Description, src.Desc, StringComparison.Ordinal))
                    { rec.Description = src.Desc; changed = true; }

                    if (kind == TableSync.XingTableType.Main)
                    {
                        if (!string.Equals(rec.Location, src.Loc, StringComparison.Ordinal))
                        { rec.Location = src.Loc; changed = true; }

                        if (!string.Equals(rec.DwgRef, src.Dwg, StringComparison.Ordinal))
                        { rec.DwgRef = src.Dwg; changed = true; }
                    }

                    matched++;
                    if (changed)
                    {
                        updated++;
                        ed.WriteMessage($"\n[CrossingManager] [U] {rec.Crossing}: grid updated from table (row {src.Row}).");
                    }
                    else
                    {
                        ed.WriteMessage($"\n[CrossingManager] [=] {rec.Crossing}: no changes needed (already matches).");
                    }
                }

                tr.Commit();
                gridCrossings.Refresh();
                _isDirty = true;

                ed.WriteMessage($"\n[CrossingManager] Match Table -> grid only (X-only): matched={matched}, updated={updated}, noMatch={noMatch}");

                gridChanged = (updated > 0);
            } // end read/lock

            // === Persist to DWG + tables if requested ===
            if (persistAfterMatch && gridChanged)
            {
                try
                {
                    // 1) Write block attributes from the grid
                    _repository.ApplyChanges(_records.ToList(), _tableSync);

                    // 2) Update MAIN/PAGE/LATLONG tables from the grid (B..E only — never col 0)
                    UpdateAllXingTablesFromGrid();

                    // 3) Reload grid from DWG (no table writes)
                    RescanRecords(applyToTables: false);

                    _isDirty = false;
                    ed.WriteMessage("\n[CrossingManager] MATCH TABLE: applied changes to DWG (blocks & tables).");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return gridChanged;
        }

        /// Read Column A strictly from the block attribute living on the cell *content*.
        /// We try both signatures seen across releases:
        ///   GetBlockAttributeValue(string tag)
        ///   GetBlockAttributeValue(ObjectId attDefId)
        /// Returns the raw value ("X1", "X02", ...) or "" if not present.
        private static string ReadXFromCellAttributeOnly(Table table, int row, Transaction tr)
        {
            if (table == null || row < 0 || row >= table.Rows.Count) return string.Empty;

            // Common attribute tags used for the bubble/index
            string[] tags = { "CROSSING", "X", "XING", "X_NO", "XNUM", "XNUMBER", "NUMBER", "INDEX", "NO", "LABEL" };

            Cell cell = null;
            try { cell = table.Cells[row, 0]; } catch { return string.Empty; }
            if (cell == null) return string.Empty;

            foreach (var content in EnumerateCellContents(cell))
            {
                var ct = content.GetType();

                // (A) GetBlockAttributeValue(string tag)
                var miStr = ct.GetMethod("GetBlockAttributeValue", new[] { typeof(string) });
                if (miStr != null)
                {
                    foreach (var tag in tags)
                    {
                        try
                        {
                            var res = miStr.Invoke(content, new object[] { tag });
                            var s = Convert.ToString(res, CultureInfo.InvariantCulture);
                            if (!string.IsNullOrWhiteSpace(s))
                                return s.Trim();
                        }
                        catch { /* try next tag */ }
                    }
                }

                // (B) GetBlockAttributeValue(ObjectId attDefId)
                var miId = ct.GetMethod("GetBlockAttributeValue", new[] { typeof(ObjectId) });
                if (miId != null)
                {
                    var btrId = TryGetBlockTableRecordIdFromContent(content);
                    if (!btrId.IsNull)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        foreach (ObjectId id in btr)
                        {
                            var ad = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                            if (ad == null) continue;

                            try
                            {
                                var res = miId.Invoke(content, new object[] { ad.ObjectId });
                                var s = Convert.ToString(res, CultureInfo.InvariantCulture);
                                if (!string.IsNullOrWhiteSpace(s))
                                    return s.Trim();
                            }
                            catch { /* try next attdef */ }
                        }
                    }
                }
            }

            // STRICT: no text/name fallback here
            return string.Empty;
        }

        /// Enumerate the cell's "Contents" collection (works across versions).
        private static IEnumerable<object> EnumerateCellContents(Cell cell)
        {
            if (cell == null) yield break;

            var contentsProp = cell.GetType().GetProperty(
                "Contents",
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

            System.Collections.IEnumerable seq = null;
            try { seq = contentsProp?.GetValue(cell, null) as System.Collections.IEnumerable; } catch { }
            if (seq == null) yield break;

            foreach (var item in seq)
                if (item != null) yield return item;
        }

        /// Pull BlockTableRecordId out of a content item (if exposed).
        private static ObjectId TryGetBlockTableRecordIdFromContent(object content)
        {
            if (content == null) return ObjectId.Null;

            var btrProp = content.GetType().GetProperty(
                "BlockTableRecordId",
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

            try
            {
                var val = btrProp?.GetValue(content, null);
                if (val is ObjectId id) return id;
            }
            catch { }

            return ObjectId.Null;
        }

        private static string GetBlockEffectiveName(BlockReference br, Transaction tr)
        {
            if (br == null || tr == null) return string.Empty;

            try
            {
                var btr = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
                return btr?.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static Dictionary<string, string> ReadAttributes(BlockReference br, Transaction tr)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (br?.AttributeCollection == null || tr == null) return values;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                if (!attId.IsValid) continue;

                AttributeReference attRef = null;
                try { attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference; }
                catch { attRef = null; }

                if (attRef == null) continue;

                values[attRef.Tag] = attRef.TextString;
            }

            return values;
        }

        private static string GetAttributeValue(IDictionary<string, string> attributes, string key)
        {
            if (attributes == null || string.IsNullOrEmpty(key)) return string.Empty;
            return attributes.TryGetValue(key, out var value) ? value : string.Empty;
        }

        private static void EnsureLayerExists(Transaction tr, Database db, string layerName)
        {
            if (tr == null || db == null || string.IsNullOrWhiteSpace(layerName)) return;

            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (lt.Has(layerName)) return;

            lt.UpgradeOpen();

            var ltr = new LayerTableRecord
            {
                Name = layerName
            };

            lt.Add(ltr);
            tr.AddNewlyCreatedDBObject(ltr, true);
        }

        private static string NormalizeXKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;

            var up = Regex.Replace(s.ToUpperInvariant(), @"\s+", "");
            // "X##" -> "X#"
            var m = Regex.Match(up, @"^X0*(\d+)$");
            if (m.Success) return "X" + m.Groups[1].Value;
            // "##"  -> "X#"
            m = Regex.Match(up, @"^0*(\d+)$");
            if (m.Success) return "X" + m.Groups[1].Value;

            return up;
        }

        private static string ReadTableCellText(Table t, int row, int col)
        {
            if (t == null || row < 0 || col < 0) return string.Empty;
            if (row >= t.Rows.Count || col >= t.Columns.Count) return string.Empty;

            try
            {
                var s = t.Cells[row, col]?.TextString ?? string.Empty;
                s = s.Trim();
                return (s == "-" || s == "—") ? string.Empty : s;
            }
            catch { return string.Empty; }
        }

        private static string ExtractXToken(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            var s = Regex.Replace(text, @"\s+", "");
            var m = Regex.Match(s, @"[Xx]0*(\d+)");
            if (m.Success) return "X" + m.Groups[1].Value;
            m = Regex.Match(s, @"^0*(\d+)$");
            if (m.Success) return "X" + m.Groups[1].Value;
            return string.Empty;
        }

        private void DeleteRowFromTables(string crossingKey)
        {
            var normalized = NormalizeXKey(crossingKey);
            var doc = _doc ?? AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (doc.LockDocument())
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    foreach (ObjectId entId in btr)
                    {
                        var table = tr.GetObject(entId, OpenMode.ForRead) as Table;
                        if (table == null) continue;

                        var kind = _tableSync.IdentifyTable(table, tr);
                        if (kind == TableSync.XingTableType.Unknown) continue;

                        table.UpgradeOpen();

                        int startRow = 0;
                        if (kind == TableSync.XingTableType.LatLong)
                        {
                            try { startRow = TableSync.FindLatLongDataStartRow(table); }
                            catch { startRow = 0; }
                            if (startRow < 0) startRow = 0;
                        }

                        for (int row = startRow; row < table.Rows.Count; row++)
                        {
                            // Column A: attribute-first key; fall back to plain text token.
                            string xRaw = ReadXFromCellAttributeOnly(table, row, tr);
                            if (string.IsNullOrWhiteSpace(xRaw))
                            {
                                var txt = table.Cells[row, 0]?.TextString ?? string.Empty;
                                xRaw = ExtractXToken(txt);
                            }

                            var key = NormalizeXKey(xRaw);
                            if (!normalized.Equals(key, StringComparison.OrdinalIgnoreCase))
                                continue;

                            // Delete the entire row (no shifting across columns).
                            table.DeleteRows(row, 1);

                            // Keep table internals and graphics up to date.
                            try { table.GenerateLayout(); } catch { }
                            try { table.RecordGraphicsModified(true); } catch { }
                            ForceRegenTable(table);

                            break; // done with this table
                        }
                    }
                }

                tr.Commit();
            }

            // Extra flush
            try { doc.Editor?.Regen(); } catch { }
            try { AcadApp.UpdateScreen(); } catch { }
        }

        // Modify your btnDelete_Click handler so it doesn’t renumber bubbles
        // and calls DeleteRowFromTables and UpdateAllXingTablesFromGrid.

        private CrossingRecord GetSelectedRecord()
        {
            if (gridCrossings.CurrentRow == null) return null;
            return gridCrossings.CurrentRow.DataBoundItem as CrossingRecord;
        }

        /// Force a visual rebuild of every recognised crossing table (MAIN / PAGE / LATLONG).
        /// <summary>
        /// Rebuilds the display lists for all recognized Crossing tables (MAIN/PAGE/LATLONG)
        /// without changing geometry. Uses non-undoable screen refreshes.
        /// </summary>
        private void ForceRegenAllCrossingTablesInDwg()
        {
            var doc = _doc ?? AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (doc.LockDocument())
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    foreach (ObjectId entId in btr)
                    {
                        var table = tr.GetObject(entId, OpenMode.ForRead) as Table;
                        if (table == null) continue;

                        var kind = _tableSync.IdentifyTable(table, tr);
                        if (kind == TableSync.XingTableType.Unknown) continue;

                        ForceRegenTable(table);
                    }
                }
                tr.Commit();
            }

            // Non-undoable repaint paths
            try { doc.Editor?.Regen(); } catch { /* ignore */ }
            try { AcadApp.UpdateScreen(); } catch { /* ignore */ }
            try { doc.SendStringToExecute("._REGENALL ", true, false, true); } catch { /* ignore */ }
        }

        /// Force AutoCAD to rebuild graphics for a single table by doing a tiny no-op transform.
        /// This mimics a user "move/unmove" which clears the transient ghost row artifact.
        /// <summary>
        /// Force a redraw of a single Table without making any geometric changes.
        /// This does NOT add anything to the Undo stack.
        /// </summary>
        private static void ForceRegenTable(Table table)
        {
            if (table == null) return;

            try
            {
                table.UpgradeOpen();

                // Make sure internal row/column cache is current
                try { table.GenerateLayout(); } catch { /* best effort */ }

                // Mark graphics “dirty” so the display list is rebuilt
                try { table.RecordGraphicsModified(true); } catch { /* older versions: no-op */ }
            }
            catch
            {
                // purely visual hygiene; ignore failures
            }
        }

        private void AddRncPolyline()
        {
            var doc = _doc ?? AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                MessageBox.Show(
                    "No active drawing is available for creating a polyline.",
                    "Crossing Manager",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            var ed = doc.Editor;
            if (ed == null)
            {
                MessageBox.Show(
                    "The active drawing does not have an editor available.",
                    "Crossing Manager",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            var polylinePrompt = new PromptEntityOptions("\nSelect boundary polyline:")
            {
                AllowNone = false
            };
            polylinePrompt.SetRejectMessage("\nEntity must be a polyline.");
            polylinePrompt.AddAllowedClass(typeof(Polyline), exactMatch: false);

            var boundaryResult = ed.GetEntity(polylinePrompt);
            if (boundaryResult.Status != PromptStatus.OK)
            {
                return;
            }

            bool created = false;
            bool insufficient = false;
            string boundaryError = null;

            using (doc.LockDocument())
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var boundary = tr.GetObject(boundaryResult.ObjectId, OpenMode.ForRead) as Polyline;
                if (boundary == null)
                {
                    boundaryError = "Selected entity is not a supported polyline.";
                }
                else if (!boundary.Closed)
                {
                    boundaryError = "Selected polyline must be closed to define a boundary.";
                }
                else
                {
                    var window = new Point3dCollection();
                    for (int i = 0; i < boundary.NumberOfVertices; i++)
                    {
                        window.Add(boundary.GetPoint3dAt(i));
                    }

                    if (window.Count < 3)
                    {
                        boundaryError = "Selected polyline does not contain enough vertices to define an area.";
                    }
                    else
                    {
                        var tolerance = new Tolerance(1e-6, 1e-6);
                        var first = window[0];
                        var last = window[window.Count - 1];
                        if (!first.IsEqualTo(last, tolerance))
                        {
                            window.Add(first);
                        }

                        var selection = ed.SelectWindowPolygon(window);
                        if (selection.Status == PromptStatus.Cancel)
                        {
                            return;
                        }

                        if (selection.Status != PromptStatus.OK || selection.Value.Count == 0)
                        {
                            insufficient = true;
                        }
                        else
                        {
                            var candidates = new List<(string Crossing, Point3d Position)>();

                            foreach (SelectedObject sel in selection.Value)
                            {
                                if (sel == null) continue;
                                if (sel.ObjectId == boundaryResult.ObjectId) continue;

                                var br = tr.GetObject(sel.ObjectId, OpenMode.ForRead) as BlockReference;
                                if (br == null) continue;

                                var effectiveName = GetBlockEffectiveName(br, tr);
                                if (!string.Equals(effectiveName, XingRepository.BlockName, StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }

                                var attributes = ReadAttributes(br, tr);
                                var crossingValue = GetAttributeValue(attributes, "CROSSING");
                                if (string.IsNullOrWhiteSpace(crossingValue))
                                {
                                    continue;
                                }

                                candidates.Add((crossingValue.Trim(), br.Position));
                            }

                            if (candidates.Count < 2)
                            {
                                insufficient = true;
                            }
                            else
                            {
                                candidates.Sort((left, right) => CrossingRecord.CompareCrossingKeys(left.Crossing, right.Crossing));

                                EnsureLayerExists(tr, doc.Database, "DEFPOINTS");

                                var owner = (BlockTableRecord)tr.GetObject(boundary.OwnerId, OpenMode.ForWrite);

                                var polyline = new Polyline();
                                polyline.SetDatabaseDefaults();
                                polyline.Layer = "DEFPOINTS";

                                var elevation = candidates[0].Position.Z;
                                polyline.Elevation = elevation;

                                for (int i = 0; i < candidates.Count; i++)
                                {
                                    var pt = candidates[i].Position;
                                    polyline.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0.0, 0.0, 0.0);
                                }

                                owner.AppendEntity(polyline);
                                tr.AddNewlyCreatedDBObject(polyline, true);

                                tr.Commit();
                                created = true;
                            }
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(boundaryError))
            {
                MessageBox.Show(
                    boundaryError,
                    "Crossing Manager",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (insufficient)
            {
                MessageBox.Show(
                    "Select at least two xing2 blocks with a CROSSING attribute inside the boundary polyline to create the polyline.",
                    "Crossing Manager",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            if (created)
            {
                try { doc.Editor?.Regen(); } catch { }
            }
        }

        private void StartRenumberCrossingCommand()
        {
            var doc = _doc ?? AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                MessageBox.Show(
                    "No active drawing is available for renumbering.",
                    "Crossing Manager",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (_isAwaitingRenumber)
            {
                MessageBox.Show(
                    "Renumber Crossing is already running. Complete the prompts in AutoCAD.",
                    "Crossing Manager",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            _isAwaitingRenumber = true;

            CommandEventHandler ended = null;
            CommandEventHandler cancelled = null;
            CommandEventHandler failed = null;

            bool IsRenumberCommand(CommandEventArgs args)
            {
                if (!_isAwaitingRenumber)
                {
                    return false;
                }

                var name = args?.GlobalCommandName;
                if (string.IsNullOrWhiteSpace(name))
                {
                    return true;
                }

                if (string.Equals(name, "RNC", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, "XINGREN", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (name.IndexOf("RENUMBER", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                if (name.IndexOf("XING", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    name.IndexOf("REN", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                return false;
            }

            void Cleanup()
            {
                doc.CommandEnded -= ended;
                doc.CommandCancelled -= cancelled;
                if (failed != null)
                {
                    try { doc.CommandFailed -= failed; } catch { /* ignore */ }
                }
                _isAwaitingRenumber = false;
            }

            void RescanOnUiThread()
            {
                void Rescan() { try { RescanRecords(); } catch { /* best effort */ } }

                if (InvokeRequired)
                {
                    try { BeginInvoke((Action)Rescan); }
                    catch { Rescan(); }
                }
                else
                {
                    Rescan();
                }
            }

            ended = (sender, args) =>
            {
                if (!IsRenumberCommand(args)) return;

                Cleanup();
                RescanOnUiThread();
            };

            cancelled = (sender, args) =>
            {
                if (!IsRenumberCommand(args)) return;

                Cleanup();
            };

            failed = (sender, args) =>
            {
                if (!IsRenumberCommand(args)) return;

                Cleanup();
                try
                {
                    MessageBox.Show(
                        "Renumber Crossing command failed.",
                        "Crossing Manager",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                catch
                {
                    // ignored
                }
            };

            doc.CommandEnded += ended;
            doc.CommandCancelled += cancelled;
            try { doc.CommandFailed += failed; }
            catch { failed = null; }

            try
            {
                doc.SendStringToExecute("RNC\n", true, false, true);
            }
            catch (Exception ex)
            {
                Cleanup();
                MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        // ===== Page & Lat/Long creation =====

        private sealed class AllPagesOption
        {
            public string DwgRef { get; set; }
            public bool IncludeAdjacent { get; set; }
            public string SelectedLocation { get; set; }
            public bool LocationEditable { get; set; }
        }

        private sealed class NaturalDwgRefComparer : IComparer<string>
        {
            public int Compare(string x, string y)
            {
                if (ReferenceEquals(x, y)) return 0;
                if (x == null) return -1;
                if (y == null) return 1;

                int ix = 0, iy = 0;
                while (ix < x.Length && iy < y.Length)
                {
                    var segX = ReadSegment(x, ref ix, out bool isNumericX);
                    var segY = ReadSegment(y, ref iy, out bool isNumericY);

                    if (isNumericX && isNumericY)
                    {
                        int cmp = CompareNumericSegments(segX, segY);
                        if (cmp != 0) return cmp;
                        continue;
                    }

                    if (isNumericX != isNumericY)
                    {
                        return isNumericX ? -1 : 1;
                    }

                    int textCmp = string.Compare(segX, segY, StringComparison.OrdinalIgnoreCase);
                    if (textCmp != 0) return textCmp;
                }

                if (ix < x.Length) return 1;
                if (iy < y.Length) return -1;
                return 0;
            }

            private static string ReadSegment(string value, ref int index, out bool isNumeric)
            {
                if (string.IsNullOrEmpty(value) || index >= value.Length)
                {
                    isNumeric = false;
                    return string.Empty;
                }

                isNumeric = char.IsDigit(value[index]);
                int start = index;
                while (index < value.Length && char.IsDigit(value[index]) == isNumeric)
                {
                    index++;
                }

                return value.Substring(start, index - start);
            }

            private static int CompareNumericSegments(string leftSegment, string rightSegment)
            {
                var left = TrimLeadingZeros(leftSegment, out int leftTrimmed);
                var right = TrimLeadingZeros(rightSegment, out int rightTrimmed);

                if (left.Length != right.Length)
                    return left.Length < right.Length ? -1 : 1;

                int cmp = string.Compare(left, right, StringComparison.Ordinal);
                if (cmp != 0) return cmp;

                return leftTrimmed.CompareTo(rightTrimmed);
            }

            private static string TrimLeadingZeros(string value, out int trimmedCount)
            {
                if (string.IsNullOrEmpty(value))
                {
                    trimmedCount = 0;
                    return string.Empty;
                }

                int i = 0;
                while (i < value.Length && value[i] == '0')
                {
                    i++;
                }

                if (i == value.Length)
                {
                    trimmedCount = Math.Max(0, value.Length - 1);
                    return "0";
                }

                trimmedCount = i;
                return value.Substring(i);
            }
        }

        private static (int Missing, int Number, string Suffix) BuildCrossingSortKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return (1, int.MaxValue, string.Empty);
            }

            var token = CrossingRecord.ParseCrossingNumber(value);
            var suffix = token.Suffix?.Trim().ToUpperInvariant() ?? string.Empty;
            return (0, token.Number, suffix);
        }

        private sealed class PageGenerationOptions
        {
            public PageGenerationOptions(string dwgRef, bool includeAdjacent, bool generateAll = false)
            {
                DwgRef = dwgRef;
                IncludeAdjacent = includeAdjacent;
                GenerateAll = generateAll;
            }

            public string DwgRef { get; }

            public bool IncludeAdjacent { get; }

            public bool GenerateAll { get; }
        }

        private void GenerateAllXingPages()
        {
            var refs = _records
                .Select(r => r.DwgRef ?? string.Empty)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s, DwgRefComparer)
                .ToList();

            if (refs.Count == 0)
            {
                MessageBox.Show("No DWG_REF values available.", "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Build per‑DWG_REF list of distinct LOCATIONs
            var locMap = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (var dr in refs)
            {
                var locs = _records
                    .Where(r => string.Equals(r.DwgRef ?? string.Empty, dr, StringComparison.OrdinalIgnoreCase))
                    .Select(r => r.Location ?? string.Empty)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                locMap[dr] = locs;
            }

            // Dialog: per DWG_REF -> IncludeAdjacent + LOCATION (if multiple)
            var options = PromptForAllPagesOptions(refs, locMap);
            if (options == null || options.Count == 0) return;

            options = options
                .Select(o => new
                {
                    Option = o,
                    EarliestCrossing = GetEarliestCrossingForDwgRef(o?.DwgRef ?? string.Empty)
                })
                .OrderBy(x => BuildCrossingSortKey(x.EarliestCrossing))
                .ThenBy(x => BuildCrossingSortKey(x.Option?.DwgRef ?? string.Empty))
                .ThenBy(x => x.Option?.DwgRef ?? string.Empty, DwgRefComparer)
                .Select(x => x.Option)
                .ToList();

            const double TableVerticalGap = 10.0;

            try
            {
                using (_doc.LockDocument())
                {
                    foreach (var opt in options)
                    {
                        // 1) Clone layout
                        string actualName;
                        var layoutId = _layoutUtils.CloneLayoutFromTemplate(
                            _doc,
                            TemplatePath,
                            GetTemplateLayoutNameForDwgRef(opt.DwgRef),
                            string.Format(CultureInfo.InvariantCulture, "XING #{0}", opt.DwgRef),
                            out actualName);

                        // 2) Heading
                        _layoutUtils.UpdatePlanHeadingText(_doc.Database, layoutId, opt.IncludeAdjacent);

                        // 3) LOCATION placeholder
                        var loc = opt.SelectedLocation ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(loc))
                        {
                            if (_layoutUtils.TryFormatMeridianLocation(loc, out var formatted))
                                loc = formatted;

                            _layoutUtils.ReplacePlaceholderText(_doc.Database, layoutId, loc);
                        }

                        // 4) Centered insertion of PAGE table
                        //    Compute center of layout, ask TableSync for table size, then offset to LL insert point.
                        Point3d center;
                        using (var tr = _doc.Database.TransactionManager.StartTransaction())
                        {
                            center = GetLayoutCenter(layoutId, tr);
                            tr.Commit();
                        }

                        var dataRowCount = _records.Count(r =>
                            string.Equals(r.DwgRef ?? string.Empty, opt.DwgRef, StringComparison.OrdinalIgnoreCase));

                        double totalW, totalH;
                        _tableSync.GetPageTableSize(dataRowCount, out totalW, out totalH);

                        var latLongRecords = _records
                            .Where(r => string.Equals(r.DwgRef ?? string.Empty, opt.DwgRef, StringComparison.OrdinalIgnoreCase))
                            .Where(HasLatLongData)
                            .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                            .ToList();

                        double latTotalW = 0.0, latTotalH = 0.0;
                        if (latLongRecords.Count > 0)
                        {
                            _tableSync.GetLatLongTableSize(latLongRecords.Count, out latTotalW, out latTotalH);
                        }

                        var insert = new Point3d(center.X - totalW / 2.0, center.Y - totalH / 2.0, 0.0);
                        var latInsert = latLongRecords.Count > 0
                            ? new Point3d(center.X - latTotalW / 2.0, insert.Y + totalH + TableVerticalGap, 0.0)
                            : Point3d.Origin;

                        using (var tr = _doc.Database.TransactionManager.StartTransaction())
                        {
                            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                            var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

                            _tableSync.CreateAndInsertPageTable(_doc.Database, tr, btr, insert, opt.DwgRef, _records);

                            if (latLongRecords.Count > 0)
                            {
                                _tableSync.CreateAndInsertLatLongTable(_doc.Database, tr, btr, latInsert, latLongRecords);
                            }

                            tr.Commit();
                        }

                        // Optional: switch to the new layout to confirm creation
                        _layoutUtils.SwitchToLayout(_doc, actualName);
                    }

                    // After creating all XING pages, reorder them by the numeric part
                    // of the layout name so that XING #1, XING #2, etc. appear in order.
                    try
                    {
                        ReorderXingLayouts();
                    }
                    catch
                    {
                        // best effort: if reordering fails, don’t abort page creation
                    }
                }
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show($"Template not found: {TemplatePath}", "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Reorders all layouts whose name begins with "XING #"
        /// by the numeric part of their name.  AutoCAD assigns new layouts
        /// arbitrary tab orders when they are created; this method ensures
        /// that the layout tabs appear in ascending numeric order
        /// (e.g., XING #1, XING #2, XING #3).  The “Model” tab (order 0)
        /// is not affected.
        /// </summary>
        private void ReorderXingLayouts()
        {
            var doc = _doc ?? AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                var list = new List<(string Name, int Number)>();

                // Collect all layout names that start with "XING #"
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layoutId = entry.Value;
                    var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                    var name = layout.LayoutName ?? string.Empty;
                    if (name.StartsWith("XING #", StringComparison.OrdinalIgnoreCase))
                    {
                        var part = name.Substring(6).Trim();
                        if (int.TryParse(part, out var num))
                        {
                            list.Add((name, num));
                        }
                    }
                }

                // Sort the list numerically by the extracted number
                list.Sort((a, b) => a.Number.CompareTo(b.Number));

                // Assign new tab orders starting from 1 (Model is always 0)
                var mgr = LayoutManager.Current;
                int tab = 1;
                foreach (var item in list)
                {
                    var id = mgr.GetLayoutId(item.Name);
                    var lay = (Layout)tr.GetObject(id, OpenMode.ForWrite);
                    lay.TabOrder = tab++;
                }

                tr.Commit();
            }
        }

        private string GetTemplateLayoutNameForDwgRef(string dwgRef)
        {
            if (string.IsNullOrWhiteSpace(dwgRef))
                return DefaultTemplateLayoutName;

            foreach (var record in _records)
            {
                if (!string.Equals(record?.DwgRef ?? string.Empty, dwgRef, StringComparison.OrdinalIgnoreCase))
                    continue;

                var description = record?.Description;
                if (!string.IsNullOrWhiteSpace(description))
                {
                    if (RailwayKeywords.Any(keyword =>
                            description.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        return CnrTemplateLayoutName;
                    }

                    if (HighwayKeywords.Any(keyword =>
                            description.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        return HighwayTemplateLayoutName;
                    }
                }

                var owner = record?.Owner;
                if (!string.IsNullOrWhiteSpace(owner) && HydroOwnerKeywords.Any(keyword =>
                        owner.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    return HydroTemplateLayoutName;
                }

                if (!string.IsNullOrWhiteSpace(description) && HydroKeywords.Any(keyword =>
                        description.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    return HydroTemplateLayoutName;
                }
            }

            return DefaultTemplateLayoutName;
        }

        private string GetEarliestCrossingForDwgRef(string dwgRef)
        {
            if (string.IsNullOrWhiteSpace(dwgRef))
                return string.Empty;

            string best = null;

            foreach (var record in _records)
            {
                if (!string.Equals(record?.DwgRef ?? string.Empty, dwgRef, StringComparison.OrdinalIgnoreCase))
                    continue;

                var crossing = record?.Crossing;
                if (string.IsNullOrWhiteSpace(crossing))
                    continue;

                if (best == null || CrossingRecord.CompareCrossingKeys(crossing, best) < 0)
                    best = crossing;
            }

            return best ?? string.Empty;
        }

        private void GenerateAllLatLongTables()
        {
            var latRecords = _records
                .Where(HasLatLongData)
                .Where(r => !string.IsNullOrWhiteSpace(r.DwgRef))
                .OrderBy(r => r.DwgRef ?? string.Empty, DwgRefComparer)
                .ThenBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();

            if (latRecords.Count == 0)
            {
                MessageBox.Show("No LAT/LONG data available to create tables.", "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var editor = _doc.Editor;
            if (editor == null)
            {
                MessageBox.Show("Unable to access the AutoCAD editor.", "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var pointRes = editor.GetPoint("\nSpecify insertion point for combined LAT/LONG table:");
            if (pointRes.Status != PromptStatus.OK)
                return;

            using (_doc.LockDocument())
            using (var tr = _doc.Database.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = null;
                try
                {
                    var layoutManager = LayoutManager.Current;
                    if (layoutManager != null)
                    {
                        var layoutDict = (DBDictionary)tr.GetObject(_doc.Database.LayoutDictionaryId, OpenMode.ForRead);
                        if (layoutDict.Contains(layoutManager.CurrentLayout))
                        {
                            var layoutId = layoutDict.GetAt(layoutManager.CurrentLayout);
                            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                            btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);
                        }
                    }
                }
                catch
                {
                    btr = null;
                }

                if (btr == null)
                {
                    btr = (BlockTableRecord)tr.GetObject(_doc.Database.CurrentSpaceId, OpenMode.ForWrite);
                }

                _tableSync.CreateAndInsertLatLongTable(_doc.Database, tr, btr, pointRes.Value, latRecords);

                tr.Commit();
            }
        }

        private List<AllPagesOption> PromptForAllPagesOptions(IList<string> dwgRefs, IDictionary<string, List<string>> locMap)
        {
            using (var dialog = new Form())
            {
                dialog.Text = "Generate All XING Pages";
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.Width = 780;
                dialog.Height = 420;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;

                var label = new Label
                {
                    Text = "Per DWG_REF: choose whether to include \"AND ADJACENT TO\" in the heading and (if needed) a LOCATION for the title block.",
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    Padding = new Padding(12, 12, 12, 8)
                };

                var grid = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    AutoGenerateColumns = false,
                    RowHeadersVisible = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    MultiSelect = false
                };

                var colRef = new DataGridViewTextBoxColumn
                {
                    HeaderText = "DWG_REF",
                    DataPropertyName = "DwgRef",
                    ReadOnly = true,
                    Width = 160
                };
                var colAdj = new DataGridViewCheckBoxColumn
                {
                    HeaderText = "Include \"AND ADJACENT TO\"",
                    DataPropertyName = "IncludeAdjacent",
                    Width = 220
                };
                var colLoc = new DataGridViewComboBoxColumn
                {
                    HeaderText = "LOCATION (if multiple)",
                    DataPropertyName = "SelectedLocation",
                    Width = 360
                };

                grid.Columns.Add(colRef);
                grid.Columns.Add(colAdj);
                grid.Columns.Add(colLoc);

                var union = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { string.Empty };
                foreach (var locs in locMap.Values)
                {
                    foreach (var loc in locs)
                    {
                        if (!string.IsNullOrWhiteSpace(loc))
                        {
                            union.Add(loc);
                        }
                    }
                }

                var orderedUnion = union
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (union.Contains(string.Empty))
                    colLoc.Items.Add(string.Empty);

                foreach (var loc in orderedUnion)
                    colLoc.Items.Add(loc);

                var binding = new BindingList<AllPagesOption>();
                foreach (var dr in dwgRefs)
                {
                    var locs = locMap.ContainsKey(dr) ? locMap[dr] : new List<string>();
                    var opt = new AllPagesOption
                    {
                        DwgRef = dr,
                        IncludeAdjacent = false,
                        SelectedLocation = (locs.Count > 0 ? locs[0] : string.Empty),
                        LocationEditable = locs.Count > 1
                    };
                    binding.Add(opt);
                }
                grid.DataSource = binding;

                void ApplyLocationCellState()
                {
                    foreach (DataGridViewRow row in grid.Rows)
                    {
                        if (row?.Cells.Count <= colLoc.Index) continue;
                        if (row.DataBoundItem is AllPagesOption opt)
                        {
                            var cell = row.Cells[colLoc.Index] as DataGridViewComboBoxCell;
                            if (cell == null) continue;

                            var readOnly = !opt.LocationEditable;
                            cell.ReadOnly = readOnly;
                            cell.DisplayStyle = readOnly
                                ? DataGridViewComboBoxDisplayStyle.Nothing
                                : DataGridViewComboBoxDisplayStyle.DropDownButton;
                            cell.Style.ForeColor = readOnly
                                ? SystemColors.GrayText
                                : grid.DefaultCellStyle.ForeColor;
                            cell.Style.BackColor = readOnly
                                ? SystemColors.Control
                                : grid.DefaultCellStyle.BackColor;
                        }
                    }
                }

                ApplyLocationCellState();
                grid.DataBindingComplete += (s, e) => ApplyLocationCellState();

                // Populate the LOCATION dropdown per row right before editing
                grid.CellBeginEdit += (s, e) =>
                {
                    if (e.RowIndex < 0) return;
                    if (!(grid.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn)) return;

                    if (!(grid.Rows[e.RowIndex].DataBoundItem is AllPagesOption opt)) return;
                    if (!opt.LocationEditable)
                    {
                        e.Cancel = true;
                        return;
                    }

                    var dr = opt.DwgRef;
                    var items = locMap.ContainsKey(dr) ? locMap[dr] : new List<string>();

                    grid.BeginInvoke(new Action(() =>
                    {
                        var editor = grid.EditingControl as DataGridViewComboBoxEditingControl;
                        if (editor != null)
                        {
                            editor.Items.Clear();
                            if (items.Count == 0) editor.Items.Add(string.Empty);
                            else foreach (var it in items) editor.Items.Add(it);
                        }
                    }));
                };

                grid.DataError += (s, e) => { e.ThrowException = false; };

                var panel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    FlowDirection = WinFormsFlowDirection.RightToLeft,
                    Height = 44,
                    Padding = new Padding(8)
                };

                var ok = new Button { Text = "Create Pages", DialogResult = DialogResult.OK, Width = 120 };
                var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 100 };
                panel.Controls.Add(ok);
                panel.Controls.Add(cancel);

                dialog.Controls.Add(grid);
                dialog.Controls.Add(label);
                dialog.Controls.Add(panel);
                dialog.AcceptButton = ok;
                dialog.CancelButton = cancel;

                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return null;

                return binding.ToList();
            }
        }

        private void GenerateXingPage()
        {
            var choices = _records
                .Select(r => r.DwgRef ?? string.Empty)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s, DwgRefComparer)
                .ToList();

            if (!choices.Any())
            {
                MessageBox.Show("No DWG_REF values available.", "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var options = PromptForPageOptions("Select DWG_REF", choices);
            if (options == null) return;

            if (options.GenerateAll)
            {
                GenerateAllXingPages();
                return;
            }

            if (string.IsNullOrEmpty(options.DwgRef)) return;

            try
            {
                string actualName;
                ObjectId layoutId;

                // --- lock for all DB changes here ---
                using (_doc.LockDocument())
                {
                    layoutId = _layoutUtils.CloneLayoutFromTemplate(
                        _doc,
                        TemplatePath,
                        GetTemplateLayoutNameForDwgRef(options.DwgRef),
                        string.Format(CultureInfo.InvariantCulture, "XING #{0}", options.DwgRef),
                        out actualName);

                    // Update heading + location while still locked
                    _layoutUtils.UpdatePlanHeadingText(_doc.Database, layoutId, options.IncludeAdjacent);

                    var locationText = BuildLocationText(options.DwgRef);
                    if (!string.IsNullOrEmpty(locationText))
                        _layoutUtils.ReplacePlaceholderText(_doc.Database, layoutId, locationText);

                    // Safe to switch layouts while locked
                    _layoutUtils.SwitchToLayout(_doc, actualName);
                }
                // --- end lock ---

                var latLongRecords = _records
                    .Where(r => string.Equals(r.DwgRef ?? string.Empty, options.DwgRef, StringComparison.OrdinalIgnoreCase))
                    .Where(HasLatLongData)
                    .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                    .ToList();

                // Get insertion point (no lock needed)
                var pointRes = _doc.Editor.GetPoint("\nSpecify insertion point for Crossing Page Table:");
                if (pointRes.Status != PromptStatus.OK) return;

                // Lock again only for the table creation
                using (_doc.LockDocument())
                using (var tr = _doc.Database.TransactionManager.StartTransaction())
                {
                    var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);
                    _tableSync.CreateAndInsertPageTable(_doc.Database, tr, btr, pointRes.Value, options.DwgRef, _records);
                    tr.Commit();
                }

                if (latLongRecords.Count > 0)
                {
                    var latPrompt = _doc.Editor.GetPoint("\nSpecify insertion point for LAT/LONG table:");
                    if (latPrompt.Status == PromptStatus.OK)
                    {
                        using (_doc.LockDocument())
                        using (var tr = _doc.Database.TransactionManager.StartTransaction())
                        {
                            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                            var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);
                            _tableSync.CreateAndInsertLatLongTable(_doc.Database, tr, btr, latPrompt.Value, latLongRecords);
                            tr.Commit();
                        }
                    }
                }
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show($"Template not found: {TemplatePath}", "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string BuildLocationText(string dwgRef)
        {
            if (string.IsNullOrWhiteSpace(dwgRef)) return string.Empty;

            var locs = _records
                .Where(r => string.Equals(r.DwgRef ?? string.Empty, dwgRef, StringComparison.OrdinalIgnoreCase))
                .Select(r => r.Location ?? string.Empty)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (locs.Count == 0) return string.Empty;

            string selected;
            if (locs.Count == 1)
            {
                selected = locs[0];
            }
            else
            {
                selected = PromptForLocationChoice(dwgRef, locs);
                if (string.IsNullOrWhiteSpace(selected)) return string.Empty;
            }

            if (_layoutUtils.TryFormatMeridianLocation(selected, out var formatted))
                return formatted;

            return selected;
        }

        private string PromptForLocationChoice(string dwgRef, IList<string> locations)
        {
            using (var dialog = new Form())
            {
                dialog.Text = $"Select LOCATION for DWG_REF {dwgRef}";
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.Width = 520;
                dialog.Height = 180;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;

                var layout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    RowCount = 2,
                    Padding = new Padding(12)
                };
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var label = new Label
                {
                    Text = "Multiple LOCATION values found. Choose one for the title block:",
                    AutoSize = true,
                    Margin = new Padding(3, 0, 3, 8)
                };

                var combo = new ComboBox
                {
                    Dock = DockStyle.Top,
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Margin = new Padding(3, 0, 3, 6)
                };
                foreach (var s in locations) combo.Items.Add(s);
                if (locations.Count > 0) combo.SelectedIndex = 0;

                layout.Controls.Add(label, 0, 0);
                layout.Controls.Add(combo, 0, 1);

                var panel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    FlowDirection = WinFormsFlowDirection.RightToLeft,
                    Height = 40,
                    Padding = new Padding(3, 3, 3, 6)
                };

                var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 80 };
                var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 80 };
                panel.Controls.Add(ok);
                panel.Controls.Add(cancel);

                dialog.Controls.Add(layout);
                dialog.Controls.Add(panel);
                dialog.AcceptButton = ok;
                dialog.CancelButton = cancel;

                var result = dialog.ShowDialog(this);
                if (result != DialogResult.OK) return null;

                return combo.SelectedItem as string;
            }
        }
        // Add this helper method anywhere inside XingForm
        private void FlushPendingGridEdits()
        {
            try { gridCrossings.EndEdit(); } catch { /* best effort */ }
            try { gridCrossings.CommitEdit(DataGridViewDataErrorContexts.Commit); } catch { }
            try { this.Validate(); } catch { } // pushes control-bound edits
            try
            {
                // In case a CurrencyManager is backing the BindingList
                var cm = BindingContext[_records] as CurrencyManager;
                cm?.EndCurrentEdit();
            }
            catch { }
        }

        // 1) Try Layout.Limits; 2) fall back to paperspace BTR extents; 3) (0,0).
        // Compute the visual center of a layout by unioning the extents of all
        // visible entities on its paperspace BlockTableRecord.
        // No use of Layout.Limits* or BlockTableRecord.GeometricExtents (not available in all versions).
        private static Point3d GetLayoutCenter(ObjectId layoutId, Transaction tr)
        {
            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
            var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

            bool haveExtents = false;
            Extents3d ex = new Extents3d();

            foreach (ObjectId id in btr)
            {
                // Only entities have GeometricExtents
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                try
                {
                    var e = ent.GeometricExtents;       // available on Entity across versions
                    if (!haveExtents)
                    {
                        ex = e;
                        haveExtents = true;
                    }
                    else
                    {
                        ex.AddPoint(e.MinPoint);
                        ex.AddPoint(e.MaxPoint);
                    }
                }
                catch
                {
                    // Some entities may not report extents in every state; ignore and continue
                }
            }

            if (haveExtents)
            {
                return new Point3d(
                    (ex.MinPoint.X + ex.MaxPoint.X) * 0.5,
                    (ex.MinPoint.Y + ex.MaxPoint.Y) * 0.5,
                    0.0);
            }

            // Fallback if the layout is empty (or all entities had no extents)
            return Point3d.Origin;
        }

        private PageGenerationOptions PromptForPageOptions(string title, IList<string> choices)
        {
            using (var dialog = new Form())
            {
                dialog.Text = title;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.Width = 360;
                dialog.Height = 210;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;

                var layout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    RowCount = 3,
                    Padding = new Padding(10)
                };
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var label = new Label
                {
                    Text = "Select DWG_REF:",
                    AutoSize = true,
                    Anchor = AnchorStyles.Left,
                    Margin = new Padding(3, 0, 3, 6)
                };

                var combo = new ComboBox
                {
                    Dock = DockStyle.Top,
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Margin = new Padding(3, 0, 3, 6)
                };
                combo.Items.Add(CreateAllPagesDisplayText);
                foreach (var choice in choices)
                    combo.Items.Add(choice);
                if (combo.Items.Count > 0) combo.SelectedIndex = 0;

                var adjacentCheckbox = new CheckBox
                {
                    Text = "Include \"AND ADJACENT TO\" in heading",
                    Checked = true,
                    AutoSize = true,
                    Anchor = AnchorStyles.Left,
                    Margin = new Padding(3, 6, 3, 0)
                };

                combo.SelectedIndexChanged += (s, e) =>
                {
                    var isAll = string.Equals(combo.SelectedItem as string, CreateAllPagesDisplayText, StringComparison.Ordinal);
                    adjacentCheckbox.Enabled = !isAll;
                    if (isAll)
                    {
                        adjacentCheckbox.Checked = true;
                    }
                };

                layout.Controls.Add(label, 0, 0);
                layout.Controls.Add(combo, 0, 1);
                layout.Controls.Add(adjacentCheckbox, 0, 2);

                var panel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    FlowDirection = WinFormsFlowDirection.RightToLeft,
                    Height = 40,
                    Padding = new Padding(3, 3, 3, 6)
                };

                var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 80 };
                var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 80 };
                panel.Controls.Add(ok);
                panel.Controls.Add(cancel);

                dialog.Controls.Add(layout);
                dialog.Controls.Add(panel);
                dialog.AcceptButton = ok;
                dialog.CancelButton = cancel;

                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return null;
                }

                var selected = combo.SelectedItem as string;
                if (string.IsNullOrEmpty(selected))
                {
                    return null;
                }

                if (string.Equals(selected, CreateAllPagesDisplayText, StringComparison.Ordinal))
                {
                    return new PageGenerationOptions(null, true, generateAll: true);
                }

                return new PageGenerationOptions(selected, adjacentCheckbox.Checked);
            }
        }

        private static bool HasLatLongData(CrossingRecord record)
        {
            if (record == null)
            {
                return false;
            }

            var hasCoordinate = !string.IsNullOrWhiteSpace(record.Lat) ||
                                !string.IsNullOrWhiteSpace(record.Long);
            if (!hasCoordinate)
            {
                return false;
            }

            var owner = record.Owner?.Trim();
            if (!string.IsNullOrEmpty(owner) && !string.Equals(owner, "-", StringComparison.Ordinal))
            {
                return false;
            }

            return true;
        }

        private void CreateOrUpdateLatLongTable()
        {
            var record = GetSelectedRecord();
            if (record == null)
            {
                MessageBox.Show("Select a crossing first.", "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Table existing = null;
            using (var tr = _doc.Database.TransactionManager.StartTransaction())
            {
                var blockTable = (BlockTable)tr.GetObject(_doc.Database.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in blockTable)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    foreach (ObjectId entId in btr)
                    {
                        if (!(tr.GetObject(entId, OpenMode.ForRead) is Table table)) continue;
                        if (_tableSync.IdentifyTable(table, tr) != TableSync.XingTableType.LatLong) continue;

                        for (var row = 1; row < table.Rows.Count; row++)
                        {
                            var text = table.Cells[row, 0].TextString ?? string.Empty;
                            if (string.Equals(text.Trim(), record.Crossing.Trim(), StringComparison.OrdinalIgnoreCase))
                            {
                                existing = table;
                                break;
                            }
                        }
                        if (existing != null) break;
                    }
                    if (existing != null) break;
                }
                tr.Commit();
            }

            if (existing != null)
            {
                _tableSync.UpdateAllTables(_doc, _records.ToList());
                return;
            }

            var pointRes = _doc.Editor.GetPoint("\nSpecify insertion point for LAT/LONG table:");
            if (pointRes.Status != PromptStatus.OK) return;

            using (_doc.LockDocument())
            using (var tr = _doc.Database.TransactionManager.StartTransaction())
            {
                var layoutManager = LayoutManager.Current;
                var layoutDict = (DBDictionary)tr.GetObject(_doc.Database.LayoutDictionaryId, OpenMode.ForRead);
                var layoutId = layoutDict.GetAt(layoutManager.CurrentLayout);
                var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);
                _tableSync.CreateAndInsertLatLongTable(_doc.Database, tr, btr, pointRes.Value, new List<CrossingRecord> { record });
                tr.Commit();
            }
        }

        private void AddLatLongFromDrawing()
        {
            var record = GetSelectedRecord();
            if (record == null)
            {
                MessageBox.Show("Select a crossing first.", "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var editor = _doc.Editor;
            if (editor == null)
            {
                MessageBox.Show("Unable to access the AutoCAD editor.", "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            EnsureModelSpaceActive();

            var zone = _currentUtmZone ?? PromptForUtmZone(editor);
            if (!zone.HasValue)
            {
                return;
            }

            SetCurrentUtmZone(zone);

            var point = PromptForPoint(editor, zone.Value);
            if (!point.HasValue)
            {
                return;
            }

            try
            {
                var latLong = ConvertUtmToLatLong(zone.Value, point.Value);
                var latString = latLong.lat.ToString("F6", CultureInfo.InvariantCulture);
                var longString = latLong.lon.ToString("F6", CultureInfo.InvariantCulture);

                record.Lat = latString;
                record.Long = longString;
                record.Zone = zone.Value.ToString(CultureInfo.InvariantCulture);
                gridCrossings.Refresh();
                _isDirty = true;

                MessageBox.Show(
                    string.Format(CultureInfo.InvariantCulture,
                        "Latitude: {0}\nLongitude: {1}", latString, longString),
                    "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmbUtmZone_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_isUpdatingZoneControl)
            {
                return;
            }

            int? selectedZone = null;
            if (cmbUtmZone.SelectedItem is string text && !string.IsNullOrWhiteSpace(text))
            {
                if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) &&
                    (parsed == 11 || parsed == 12))
                {
                    selectedZone = parsed;
                }
            }

            SetCurrentUtmZone(selectedZone);
        }

        private void SetCurrentUtmZone(int? zone)
        {
            if (Nullable.Equals(_currentUtmZone, zone))
            {
                UpdateZoneControlFromState();
                return;
            }

            _currentUtmZone = zone;
            UpdateZoneControlFromState();
        }

        private void UpdateZoneControlFromState()
        {
            if (cmbUtmZone == null)
            {
                return;
            }

            _isUpdatingZoneControl = true;
            try
            {
                if (_currentUtmZone.HasValue)
                {
                    var text = _currentUtmZone.Value.ToString(CultureInfo.InvariantCulture);
                    if (cmbUtmZone.Items.IndexOf(text) < 0)
                    {
                        cmbUtmZone.Items.Add(text);
                    }

                    var index = cmbUtmZone.Items.IndexOf(text);
                    cmbUtmZone.SelectedIndex = index;
                }
                else
                {
                    cmbUtmZone.SelectedIndex = -1;
                }
            }
            finally
            {
                _isUpdatingZoneControl = false;
            }
        }

        private void UpdateZoneSelectionFromRecords()
        {
            if (_records == null)
            {
                UpdateZoneControlFromState();
                return;
            }

            if (_currentUtmZone.HasValue)
            {
                UpdateZoneControlFromState();
                return;
            }

            var zones = _records
                .Select(r => r?.Zone)
                .Where(z => !string.IsNullOrWhiteSpace(z))
                .Select(z => z.Trim())
                .Select(z =>
                {
                    if (int.TryParse(z, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
                    {
                        return (int?)parsed;
                    }

                    return null;
                })
                .Where(z => z.HasValue)
                .Select(z => z.Value)
                .Distinct()
                .ToList();

            if (zones.Count == 1)
            {
                SetCurrentUtmZone(zones[0]);
            }
            else
            {
                UpdateZoneControlFromState();
            }
        }

        private void EnsureModelSpaceActive()
        {
            try
            {
                using (_doc.LockDocument())
                {
                    var layoutManager = LayoutManager.Current;
                    if (layoutManager != null &&
                        !string.Equals(layoutManager.CurrentLayout, "Model", StringComparison.OrdinalIgnoreCase))
                    {
                        layoutManager.CurrentLayout = "Model";
                    }
                }
            }
            catch
            {
                // Best effort: ignore layout switching failures.
            }
        }

        private static int? PromptForUtmZone(Editor editor)
        {
            var options = new PromptIntegerOptions("\nEnter UTM zone (11 or 12):")
            {
                AllowNone = false,
                AllowNegative = false,
                AllowZero = false,
                LowerLimit = 11,
                UpperLimit = 12
            };

            var result = editor.GetInteger(options);
            if (result.Status != PromptStatus.OK)
            {
                return null;
            }

            if (result.Value != 11 && result.Value != 12)
            {
                editor.WriteMessage("\n** Invalid zone – enter 11 or 12. **");
                return null;
            }

            return result.Value;
        }

        private static Point3d? PromptForPoint(Editor editor, int zone)
        {
            var prompt = string.Format(CultureInfo.InvariantCulture, "\nSelect point in UTM83-{0}:", zone);
            var pointOptions = new PromptPointOptions(prompt)
            {
                AllowNone = false
            };

            var pointResult = editor.GetPoint(pointOptions);
            if (pointResult.Status != PromptStatus.OK)
            {
                return null;
            }

            return pointResult.Value;
        }

        private static (double lat, double lon) ConvertUtmToLatLong(int zone, Point3d point)
        {
            if (zone != 11 && zone != 12)
            {
                throw new ArgumentOutOfRangeException(nameof(zone), "Zone must be 11 or 12.");
            }

            const double k0 = 0.9996;
            const double a = 6378137.0; // GRS80/WGS84 semi-major axis
            const double eccSquared = 0.00669438; // GRS80/WGS84 eccentricity squared

            double x = point.X - 500000.0;
            double y = point.Y;

            double eccPrimeSquared = eccSquared / (1 - eccSquared);
            double eccSquared2 = eccSquared * eccSquared;
            double eccSquared3 = eccSquared2 * eccSquared;
            double eccSquared4 = eccSquared3 * eccSquared;

            double e1 = (1 - Math.Sqrt(1 - eccSquared)) / (1 + Math.Sqrt(1 - eccSquared));
            double e1Squared = e1 * e1;
            double e1Cubed = e1Squared * e1;
            double e1Fourth = e1Cubed * e1;

            double mu = y / (k0 * a * (1 - eccSquared / 4 - 3 * eccSquared2 / 64 - 5 * eccSquared3 / 256));
            double phi1Rad = mu
                + (3 * e1 / 2 - 27 * e1Cubed / 32) * Math.Sin(2 * mu)
                + (21 * e1Squared / 16 - 55 * e1Fourth / 32) * Math.Sin(4 * mu)
                + (151 * e1Cubed / 96) * Math.Sin(6 * mu)
                + (1097 * e1Fourth / 512) * Math.Sin(8 * mu);

            double sinPhi1 = Math.Sin(phi1Rad);
            double cosPhi1 = Math.Cos(phi1Rad);
            double tanPhi1 = Math.Tan(phi1Rad);

            double n1 = a / Math.Sqrt(1 - eccSquared * sinPhi1 * sinPhi1);
            double t1 = tanPhi1 * tanPhi1;
            double c1 = eccPrimeSquared * cosPhi1 * cosPhi1;
            double r1 = a * (1 - eccSquared) / Math.Pow(1 - eccSquared * sinPhi1 * sinPhi1, 1.5);
            double d = x / (n1 * k0);

            double d2 = d * d;
            double d3 = d2 * d;
            double d4 = d2 * d2;
            double d5 = d4 * d;
            double d6 = d5 * d;
            double c1Squared = c1 * c1;
            double t1Squared = t1 * t1;

            double latRad = phi1Rad - (n1 * tanPhi1 / r1) *
                (d2 / 2 - (5 + 3 * t1 + 10 * c1 - 4 * c1Squared - 9 * eccPrimeSquared) * d4 / 24
                 + (61 + 90 * t1 + 298 * c1 + 45 * t1Squared - 252 * eccPrimeSquared - 3 * c1Squared) * d6 / 720);

            double lonRad = (d - (1 + 2 * t1 + c1) * d3 / 6
                + (5 - 2 * c1 + 28 * t1 - 3 * c1Squared + 8 * eccPrimeSquared + 24 * t1Squared) * d5 / 120) / cosPhi1;

            double longOrigin = (zone - 1) * 6 - 180 + 3;
            double lat = latRad * (180.0 / Math.PI);
            double lon = longOrigin + lonRad * (180.0 / Math.PI);

            return (lat, lon);
        }
    }
}
