/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////

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
using System.Reflection;
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

        // Tracks pending X# resequencing (old -> new).
        // This is used to keep layout-variant tables (that aren't recognized
        // by the standard table classifier) in sync after Move/Drag renumbering.
        private Dictionary<string, string> _pendingRenumberMap;

        // Optional UI help (WinForms ToolTip). Note: some AutoCAD palette hosts may
        // suppress tooltips; safe to keep even if they don't show.
        private ToolTip _toolTip;

        // Drag/drop row reordering state
        private bool _rowDragPending;
        private Point _rowDragStart;
        private int _rowDragSourceRow = -1;

        // DataGridView can collapse multi-selection on mouse down. Keep the last multi-selection
        // so drag can still move the whole block when the user starts dragging a selected row.
        private readonly List<int> _rowDragSourceIndices = new List<int>();
        private readonly List<int> _lastMultiRowSelectionIndices = new List<int>();

        // Anchor row used for SHIFT-click range selection in the grid.
        private int _rowSelectionAnchorRow = -1;
        private bool _restoreMultiSelectionOnDragStart;

        private const string TemplatePath = @"M:\Drafting\_CURRENT TEMPLATES\Compass_Main.dwt";
        private const string DefaultTemplateLayoutName = "X";
        private const string HydroTemplateLayoutName = "H2O-PROFILE";
        private const string CnrTemplateLayoutName = "CNR";
        private const string HighwayTemplateLayoutName = "HWY-PROFILE";
        private const string HydroProfileDescriptionToken = "P/L R/W";
        private static readonly string[] HydroKeywords = { "Watercourse", "Creek", "River" };
        private static readonly string[] HydroOwnerKeywords =
        {
            "Nova",
            "Alliance",
            "Pembina",
            "TCPL",
            "Ovintiv",
            "PGI",
            "NGTL"
        };
        private static readonly string[] RailwayKeywords = { "Railway" };
        private static readonly string[] HighwayKeywords = { "Highway", "Hwy" };
        private const string OtherLatLongTableTitle = "OTHER CROSSING INFORMATION";
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
            SetupTooltips();
            UpdateZoneControlFromState();
        }

        // ===== Public entry points called by your commands =====

        public void LoadData()
        {
            // Intentionally left blank to avoid automatically scanning on load.
        }

        // Plain scan refreshes the grid/duplicate UI only (no DWG writes until Apply to Drawing).
        public void RescanData() => RescanRecords(applyToTables: false);

        public void ApplyToDrawing() => ApplyChangesToDrawing();

        public void GenerateXingPageFromCommand() => GenerateXingPage();

        public void GenerateAllXingPagesFromCommand() => GenerateAllXingPages();

        public void GenerateAllLatLongTablesFromCommand() => GenerateWaterLatLongTables();

        public void GenerateOtherLatLongTablesFromCommand() => GenerateOtherLatLongTables();

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

            ConfigureGridForRowReorder();
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

        // ===== Grid row reorder (Move Up/Down + drag/drop) =====

        private const string RowDragDataFormat = "XingManager.RowIndexes";

        private void ConfigureGridForRowReorder()
        {
            if (gridCrossings == null) return;

            // Allow shift-select ranges + multi-row drag.
            gridCrossings.MultiSelect = true;
            gridCrossings.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            gridCrossings.AllowDrop = true;

            // In the AutoCAD palette host, single-click editing can interfere with Shift+Click selection.
            // Require a deliberate action (double-click or F2) to begin editing.
            gridCrossings.EditMode = DataGridViewEditMode.EditProgrammatically;

            // DataGridView's default Ctrl+C copies the full row when FullRowSelect is enabled.
            // Users typically want the current cell's value, so we take over Ctrl+C.
            gridCrossings.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;

            // Prevent sort/header clicks from reordering behind our backs.
            foreach (DataGridViewColumn col in gridCrossings.Columns)
            {
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            // Hook once (defensive)
            gridCrossings.MouseDown -= GridCrossingsOnMouseDown_ForDrag;
            gridCrossings.MouseMove -= GridCrossingsOnMouseMove_ForDrag;
            gridCrossings.DragOver -= GridCrossingsOnDragOver_ForDrag;
            gridCrossings.DragDrop -= GridCrossingsOnDragDrop_ForDrag;
            gridCrossings.KeyDown -= GridCrossingsOnKeyDown_ForCopy;
            gridCrossings.CellDoubleClick -= GridCrossingsOnCellDoubleClick_BeginEdit;
            gridCrossings.SelectionChanged -= GridCrossingsOnSelectionChanged_ForDrag;

            gridCrossings.MouseDown += GridCrossingsOnMouseDown_ForDrag;
            gridCrossings.MouseMove += GridCrossingsOnMouseMove_ForDrag;
            gridCrossings.DragOver += GridCrossingsOnDragOver_ForDrag;
            gridCrossings.DragDrop += GridCrossingsOnDragDrop_ForDrag;
            gridCrossings.KeyDown += GridCrossingsOnKeyDown_ForCopy;
            gridCrossings.CellDoubleClick += GridCrossingsOnCellDoubleClick_BeginEdit;
            gridCrossings.SelectionChanged += GridCrossingsOnSelectionChanged_ForDrag;
        }

        private void GridCrossingsOnSelectionChanged_ForDrag(object sender, EventArgs e)
        {
            // Cache last multi-selection so a mouse-down (which can collapse the selection)
            // can still drag the full block.
            try
            {
                var selected = GetSelectedRowIndices();
                if (selected.Count >= 2)
                {
                    _lastMultiRowSelectionIndices.Clear();
                    _lastMultiRowSelectionIndices.AddRange(selected);
                }
                else if (selected.Count == 0)
                {
                    _lastMultiRowSelectionIndices.Clear();
                }
            }
            catch
            {
                // ignore
            }
        }

        private void GridCrossingsOnKeyDown_ForCopy(object sender, KeyEventArgs e)
        {
            if (e == null) return;
            if (gridCrossings == null) return;

            // Start editing intentionally (F2) - we use EditProgrammatically.
            if (!e.Control && e.KeyCode == Keys.F2)
            {
                try
                {
                    var c = gridCrossings.CurrentCell;
                    if (c != null && !c.ReadOnly)
                    {
                        gridCrossings.BeginEdit(true);
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                    }
                }
                catch
                {
                    // ignore
                }
                return;
            }

            // Copy only the current cell (not the whole row)
            if (!e.Control || e.KeyCode != Keys.C) return;

            var cell = gridCrossings.CurrentCell;
            if (cell == null) return;

            try
            {
                var text = Convert.ToString(cell.EditedFormattedValue ?? cell.FormattedValue ?? cell.Value ?? string.Empty, CultureInfo.InvariantCulture);
                Clipboard.SetText(text ?? string.Empty);
            }
            catch
            {
                // Ignore clipboard errors (can happen depending on AutoCAD/palette host).
            }

            e.Handled = true;
            e.SuppressKeyPress = true;
        }

        private void GridCrossingsOnCellDoubleClick_BeginEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (gridCrossings == null) return;
            if (e == null) return;
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            try
            {
                var row = gridCrossings.Rows[e.RowIndex];
                if (row == null) return;
                var cell = row.Cells[e.ColumnIndex];
                if (cell == null || cell.ReadOnly) return;

                gridCrossings.CurrentCell = cell;
                gridCrossings.BeginEdit(true);
            }
            catch
            {
                // ignore
            }
        }

        private void GridCrossingsOnMouseDown_ForDrag(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                _rowDragPending = false;
                return;
            }

            var hit = gridCrossings.HitTest(e.X, e.Y);
            if (hit.RowIndex < 0)
            {
                _rowDragPending = false;
                return;
            }

            int clickedRow = hit.RowIndex;
            bool ctrl = (Control.ModifierKeys & Keys.Control) == Keys.Control;
            bool shift = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;

            // SHIFT-click: force a contiguous range selection. Some palette hosts interfere with the
            // DataGridView default SHIFT behavior, so we enforce it and do it on BeginInvoke so our
            // selection isn't immediately overwritten by the control's internal click logic.
            if (shift && !ctrl)
            {
                int anchorRow = _rowSelectionAnchorRow;
                if (anchorRow < 0)
                    anchorRow = gridCrossings.CurrentCell != null ? gridCrossings.CurrentCell.RowIndex : clickedRow;

                int col = hit.ColumnIndex;
                if (col < 0) col = 0;
                if (gridCrossings.Columns.Count > 0)
                    col = Math.Min(col, gridCrossings.Columns.Count - 1);
                else
                    col = 0;

                int start = Math.Min(anchorRow, clickedRow);
                int end = Math.Max(anchorRow, clickedRow);
                var range = Enumerable.Range(start, end - start + 1).ToList();

                BeginInvoke((Action)(() =>
                {
                    try
                    {
                        gridCrossings.ClearSelection();
                        for (int i = start; i <= end; i++)
                            gridCrossings.Rows[i].Selected = true;

                        if (gridCrossings.Rows[clickedRow].Cells.Count > 0 && gridCrossings.Columns.Count > 0)
                            gridCrossings.CurrentCell = gridCrossings.Rows[clickedRow].Cells[col];
                    }
                    catch
                    {
                        // ignore
                    }

                    _rowDragSourceIndices.Clear();
                    _rowDragSourceIndices.AddRange(range);

                    _lastMultiRowSelectionIndices.Clear();
                    _lastMultiRowSelectionIndices.AddRange(range);
                }));

                // Selecting with SHIFT should not arm a drag operation.
                _rowDragPending = false;
                _rowDragSourceRow = -1;
                return;
            }

            // CTRL-click should behave like normal DataGridView selection (add/remove). Don't arm drag.
            if (ctrl)
            {
                _rowDragPending = false;
                return;
            }

            // No modifiers: update the anchor for future SHIFT selection.
            _rowSelectionAnchorRow = clickedRow;

            bool clickedRowAlreadySelected = gridCrossings.Rows[clickedRow].Selected;
            bool hasMultiSelection = gridCrossings.SelectedRows != null && gridCrossings.SelectedRows.Count >= 2;

            // If user already has a multi-selection, and clicks on one of the selected rows,
            // treat it as intent to drag that selection (not to collapse selection to a single row).
            if (clickedRowAlreadySelected && hasMultiSelection)
            {
                _restoreMultiSelectionOnDragStart = true;

                // Capture selection now; if the grid collapses it later, we can restore on drag.
                var selected = GetSelectedRowIndices();
                if (selected.Count == 0 && _lastMultiRowSelectionIndices.Count >= 2)
                    selected.AddRange(_lastMultiRowSelectionIndices);

                if (selected.Count >= 2)
                {
                    _lastMultiRowSelectionIndices.Clear();
                    _lastMultiRowSelectionIndices.AddRange(selected);
                }

                _rowDragSourceRow = clickedRow;
                _rowDragSourceIndices.Clear();
                _rowDragSourceIndices.AddRange(selected);

                _rowDragStart = e.Location;
                _rowDragPending = true;
                return;
            }

            // Ensure the clicked row becomes selected (single-row selection by default).
            if (!clickedRowAlreadySelected)
            {
                gridCrossings.ClearSelection();
                gridCrossings.Rows[clickedRow].Selected = true;
            }

            // Ensure a CurrentCell exists at the click point.
            int colIndex = hit.ColumnIndex;
            if (colIndex < 0) colIndex = 0;
            if (gridCrossings.Columns.Count > 0)
                colIndex = Math.Min(colIndex, gridCrossings.Columns.Count - 1);
            else
                colIndex = 0;

            if (gridCrossings.Rows[clickedRow].Cells.Count > 0 && gridCrossings.Columns.Count > 0)
                gridCrossings.CurrentCell = gridCrossings.Rows[clickedRow].Cells[colIndex];

            _rowDragSourceRow = clickedRow;
            _rowDragSourceIndices.Clear();
            _rowDragSourceIndices.Add(clickedRow);

            _rowDragStart = e.Location;
            _rowDragPending = true;
            _restoreMultiSelectionOnDragStart = false;
        }


        private void GridCrossingsOnMouseMove_ForDrag(object sender, MouseEventArgs e)
        {
            if (!_rowDragPending) return;
            if (e.Button != MouseButtons.Left) return;
            if (gridCrossings == null) return;
            if (gridCrossings.IsCurrentCellInEditMode) return;

            var dragRect = new Rectangle(
                _rowDragStart.X - SystemInformation.DragSize.Width / 2,
                _rowDragStart.Y - SystemInformation.DragSize.Height / 2,
                SystemInformation.DragSize.Width,
                SystemInformation.DragSize.Height);

            if (dragRect.Contains(e.Location))
                return;

            _rowDragPending = false;

            var rows = _rowDragSourceIndices.Count > 0
                ? new List<int>(_rowDragSourceIndices)
                : GetSelectedRowIndices();

            rows = rows
                .Distinct()
                .Where(i => i >= 0 && i < gridCrossings.Rows.Count)
                .OrderBy(i => i)
                .ToList();

            if (rows.Count == 0) return;

            // If the grid collapsed multi-selection on click, restore it right before dragging so
            // the user sees the full block being moved.
            if (_restoreMultiSelectionOnDragStart)
            {
                RestoreRowSelection(rows, _rowDragSourceRow);
                _restoreMultiSelectionOnDragStart = false;
            }

            var data = new DataObject();
            data.SetData(RowDragDataFormat, rows.ToArray());
            gridCrossings.DoDragDrop(data, DragDropEffects.Move);
        }

        private void GridCrossingsOnDragOver_ForDrag(object sender, DragEventArgs e)
        {
            if (e.Data != null && e.Data.GetDataPresent(RowDragDataFormat))
                e.Effect = DragDropEffects.Move;
            else
                e.Effect = DragDropEffects.None;
        }

        private void GridCrossingsOnDragDrop_ForDrag(object sender, DragEventArgs e)
        {
            if (gridCrossings == null) return;
            if (e.Data == null || !e.Data.GetDataPresent(RowDragDataFormat)) return;

            var src = e.Data.GetData(RowDragDataFormat) as int[];
            if (src == null || src.Length == 0) return;

            var clientPoint = gridCrossings.PointToClient(new Point(e.X, e.Y));
            var hit = gridCrossings.HitTest(clientPoint.X, clientPoint.Y);

            int targetIndex = hit.RowIndex;
            if (targetIndex < 0)
                targetIndex = _records?.Count ?? 0; // drop after last row

            MoveRowsToIndex(src, targetIndex);
        }

        

        private void RestoreRowSelection(IList<int> rowIndices, int anchorRow)
        {
            if (gridCrossings == null) return;
            if (rowIndices == null || rowIndices.Count == 0) return;

            gridCrossings.ClearSelection();

            foreach (var idx in rowIndices.Distinct())
            {
                if (idx >= 0 && idx < gridCrossings.Rows.Count)
                    gridCrossings.Rows[idx].Selected = true;
            }

            if (anchorRow >= 0 && anchorRow < gridCrossings.Rows.Count && gridCrossings.Columns.Count > 0)
            {
                int colIndex = 0;
                if (gridCrossings.CurrentCell != null)
                    colIndex = Math.Max(0, Math.Min(gridCrossings.CurrentCell.ColumnIndex, gridCrossings.Columns.Count - 1));

                gridCrossings.CurrentCell = gridCrossings.Rows[anchorRow].Cells[colIndex];
            }
        }

        private void MoveRowsToIndex(IList<int> sourceRowIndexes, int targetIndex)
        {
            if (_records == null || _records.Count == 0) return;
            if (sourceRowIndexes == null || sourceRowIndexes.Count == 0) return;

            var src = sourceRowIndexes
                .Where(i => i >= 0 && i < _records.Count)
                .Distinct()
                .OrderBy(i => i)
                .ToList();

            if (src.Count == 0) return;

            // Allow insert at end.
            if (targetIndex < 0) targetIndex = 0;
            if (targetIndex > _records.Count) targetIndex = _records.Count;

            // Adjust insertion index after removing rows that occur before it.
            int removedBefore = src.Count(i => i < targetIndex);
            int insertIndex = targetIndex - removedBefore;

            int newCount = _records.Count - src.Count;
            if (insertIndex < 0) insertIndex = 0;
            if (insertIndex > newCount) insertIndex = newCount;

            var movedItems = src.Select(i => _records[i]).ToList();

            // Remove from bottom up.
            for (int i = src.Count - 1; i >= 0; i--)
                _records.RemoveAt(src[i]);

            for (int i = 0; i < movedItems.Count; i++)
                _records.Insert(insertIndex + i, movedItems[i]);

            // Moving rows implies resequencing X#s (user request).
            ResequenceCrossingsFromGridOrder();

            SelectRowsRange(insertIndex, movedItems.Count);
        }

        private void SelectRowsRange(int startIndex, int count)
        {
            if (gridCrossings == null) return;
            if (gridCrossings.Rows.Count == 0) return;

            gridCrossings.ClearSelection();

            for (int i = 0; i < count; i++)
            {
                int r = startIndex + i;
                if (r >= 0 && r < gridCrossings.Rows.Count)
                    gridCrossings.Rows[r].Selected = true;
            }

            int focus = Math.Min(Math.Max(startIndex, 0), gridCrossings.Rows.Count - 1);
            if (gridCrossings.Columns.Count > 0)
                gridCrossings.CurrentCell = gridCrossings.Rows[focus].Cells[0];
        }

        private void SelectRowsForRecords(ISet<CrossingRecord> selectedRecords)
        {
            if (gridCrossings == null || selectedRecords == null || selectedRecords.Count == 0)
            {
                return;
            }

            gridCrossings.ClearSelection();
            var firstRowIndex = -1;

            for (var i = 0; i < gridCrossings.Rows.Count; i++)
            {
                var row = gridCrossings.Rows[i];
                var rec = row?.DataBoundItem as CrossingRecord;
                if (rec != null && selectedRecords.Contains(rec))
                {
                    row.Selected = true;
                    if (firstRowIndex < 0)
                    {
                        firstRowIndex = i;
                    }
                }
            }

            if (firstRowIndex >= 0 && gridCrossings.Columns.Count > 0)
            {
                try
                {
                    gridCrossings.CurrentCell = gridCrossings.Rows[firstRowIndex].Cells[0];
                }
                catch
                {
                    // ignore
                }
            }
        }


        private int GetPrimarySelectedRowIndex()
        {
            if (gridCrossings == null) return -1;
            if (gridCrossings.CurrentCell != null)
                return gridCrossings.CurrentCell.RowIndex;

            if (gridCrossings.SelectedRows.Count > 0)
                return gridCrossings.SelectedRows[0].Index;

            if (gridCrossings.SelectedCells.Count > 0)
                return gridCrossings.SelectedCells[0].RowIndex;

            return -1;
        }

        private void MoveSelectedCrossingRow(int direction)
        {
            if (_records == null || _records.Count == 0)
            {
                return;
            }

            if (direction == 0)
            {
                return;
            }

            // Support multi-select moves (Shift/Ctrl selection). We move each selected row one step
            // in the requested direction, keeping the relative order of selected rows.
            var selectedIndices = GetSelectedRowIndices()
                .Where(i => i >= 0 && i < _records.Count)
                .Distinct()
                .OrderBy(i => i)
                .ToList();

            if (selectedIndices.Count == 0)
            {
                var idx = GetPrimarySelectedRowIndex();
                if (idx < 0)
                {
                    return;
                }

                selectedIndices.Add(idx);
            }

            var selectedRecords = new HashSet<CrossingRecord>(selectedIndices.Select(i => _records[i]));

            if (selectedIndices.Count == 1)
            {
                var idx = selectedIndices[0];
                var newIdx = idx + direction;
                if (newIdx < 0 || newIdx >= _records.Count)
                {
                    return;
                }

                var item = _records[idx];
                _records.RemoveAt(idx);
                _records.Insert(newIdx, item);
            }
            else
            {
                if (direction < 0)
                {
                    // Move up: walk top->bottom swapping a selected row with the row above when the above row is not selected.
                    for (var i = 1; i < _records.Count; i++)
                    {
                        if (selectedRecords.Contains(_records[i]) && !selectedRecords.Contains(_records[i - 1]))
                        {
                            var tmp = _records[i - 1];
                            _records[i - 1] = _records[i];
                            _records[i] = tmp;
                        }
                    }
                }
                else
                {
                    // Move down: walk bottom->top swapping a selected row with the row below when the below row is not selected.
                    for (var i = _records.Count - 2; i >= 0; i--)
                    {
                        if (selectedRecords.Contains(_records[i]) && !selectedRecords.Contains(_records[i + 1]))
                        {
                            var tmp = _records[i + 1];
                            _records[i + 1] = _records[i];
                            _records[i] = tmp;
                        }
                    }
                }
            }

            ResequenceCrossingsFromGridOrder();

            // Keep the same records selected after the move.
            SelectRowsForRecords(selectedRecords);
        }


        private void ResequenceCrossingsFromGridOrder()
        {
            if (_records == null || _records.Count == 0)
                return;

            // Map current X -> new X (X1..Xn).
            var currentToNew = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < _records.Count; i++)
            {
                string oldKey = NormalizeXKey(_records[i].Crossing);
                if (string.IsNullOrWhiteSpace(oldKey))
                    continue;

                string newKey = $"X{i + 1}";
                if (!currentToNew.ContainsKey(oldKey))
                    currentToNew[oldKey] = newKey;
            }

            // Compose into the pending map (originalDrawingKey -> latestKey).
            if (_pendingRenumberMap == null)
            {
                _pendingRenumberMap = new Dictionary<string, string>(currentToNew, StringComparer.OrdinalIgnoreCase);
            }
            else
            {
                // Update existing originals through the new mapping.
                var originals = _pendingRenumberMap.Keys.ToList();
                foreach (var orig in originals)
                {
                    var intermediate = _pendingRenumberMap[orig];
                    if (!string.IsNullOrWhiteSpace(intermediate) && currentToNew.TryGetValue(intermediate, out var finalKey))
                    {
                        _pendingRenumberMap[orig] = finalKey;
                    }
                    else if (currentToNew.TryGetValue(orig, out var directKey))
                    {
                        _pendingRenumberMap[orig] = directKey;
                    }
                }

                // Add any brand-new keys that didn't exist in the original map.
                foreach (var kvp in currentToNew)
                {
                    if (!_pendingRenumberMap.ContainsKey(kvp.Key))
                        _pendingRenumberMap[kvp.Key] = kvp.Value;
                }
            }

            // Apply new X#s in list order.
            for (int i = 0; i < _records.Count; i++)
            {
                _records[i].Crossing = $"X{i + 1}";
            }

            _isDirty = true;
            gridCrossings?.Refresh();
        }

        // ===== Tooltips =====

        private void SetupTooltips()
        {
            try
            {
                _toolTip = new ToolTip
                {
                    AutomaticDelay = 250,
                    AutoPopDelay = 20000,
                    InitialDelay = 250,
                    ReshowDelay = 100,
                    ShowAlways = true
                };

                _toolTip.SetToolTip(btnRescan, "Scan the drawing for crossing blocks and load them into the list.");
                _toolTip.SetToolTip(btnApply, "Write the current list values back to the drawing (blocks + tables).");
                _toolTip.SetToolTip(btnDelete, "Delete the selected crossing (blocks + matching table rows).");

                if (btnMoveUp != null) _toolTip.SetToolTip(btnMoveUp, "Move the selected row up (X#s resequence).");
                if (btnMoveDown != null) _toolTip.SetToolTip(btnMoveDown, "Move the selected row down (X#s resequence).");
                if (btnRenumberFromList != null) _toolTip.SetToolTip(btnRenumberFromList, "Resequence X#s based on current list order (X1..Xn).");

                _toolTip.SetToolTip(btnRenumber, "Run the RNC command (uses RNC Path if set) then rescan.");
                _toolTip.SetToolTip(btnAddRncPolyline, "Create/choose an RNC Path polyline used by the RNC command.");
                _toolTip.SetToolTip(btnMatchTable, "Use an existing table as the source of truth for non-X fields and push those into the blocks.");

                _toolTip.SetToolTip(btnGeneratePage, "Generate a XING PAGE table for the current list.");
                _toolTip.SetToolTip(btnGenerateAllLatLongTables, "Generate/update WATER LAT/LONG tables.");
                _toolTip.SetToolTip(btnGenerateOtherLatLongTables, "Generate/update OTHER LAT/LONG tables.");
                _toolTip.SetToolTip(btnLatLong, "Create a LAT/LONG table for the selected crossing.");
                _toolTip.SetToolTip(btnAddLatLong, "Pick points to fill LAT/LONG values for selected crossing.");

                _toolTip.SetToolTip(btnExport, "Export the list to CSV/JSON.");
                _toolTip.SetToolTip(btnImport, "Import a list from CSV/JSON.");
                _toolTip.SetToolTip(cmbUtmZone, "UTM Zone used for lat/long conversions.");
            }
            catch
            {
                // ToolTip can be finicky inside some AutoCAD palette hosts; safe to ignore.
            }
        }

        // ===== Pending renumber map -> any tables (layout variants) =====

        private void ApplyPendingRenumberMapToAnyTables()
        {
            if (_pendingRenumberMap == null || _pendingRenumberMap.Count == 0)
                return;

            if (_doc == null)
                return;

            using (_doc.LockDocument())
            using (var tr = _doc.Database.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(_doc.Database.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    foreach (ObjectId entId in btr)
                    {
                        if (entId.ObjectClass?.DxfName != "ACAD_TABLE")
                            continue;

                        Table table = null;
                        try { table = (Table)tr.GetObject(entId, OpenMode.ForWrite); }
                        catch { continue; }

                        if (table == null || table.Rows == null || table.Columns == null)
                            continue;

                        bool tableTouched = false;

                        // We assume the crossing bubble is in column 0 for these legacy/layout tables.
                        int col = 0;
                        if (table.Columns.Count <= col)
                            continue;

                        for (int r = 0; r < table.Rows.Count; r++)
                        {
                            string raw = null;
                            try { raw = TableSync.ResolveCrossingKey(table, r, col); }
                            catch { raw = null; }

                            var oldKey = NormalizeXKey(raw);
                            if (string.IsNullOrWhiteSpace(oldKey))
                                continue;

                            if (!_pendingRenumberMap.TryGetValue(oldKey, out var newKey))
                                continue;

                            if (string.Equals(oldKey, newKey, StringComparison.OrdinalIgnoreCase))
                                continue;

                            if (TrySetCrossingCellValue(table, r, col, newKey))
                                tableTouched = true;
                        }

                        if (tableTouched)
                        {
                            try { table.RecordGraphicsModified(true); } catch { }
                        }
                    }
                }

                tr.Commit();
            }
        }

        private static readonly string[] CrossingBubbleAttributeTags =
        {
            "CROSSING", "XING", "X_NO", "XING_NO", "XING#", "XINGNUM", "XING_NUM", "X", "CROSSING_ID", "CROSSINGID"
        };

        private static bool TrySetCrossingCellValue(Table table, int row, int col, string crossingText)
        {
            if (table == null) return false;

            try
            {
                var cell = table.Cells[row, col];
                if (cell == null) return false;

                // Detect whether this cell contains a block (bubble) so we don't accidentally add text content on top.
                var blockContentIndexes = new List<int>();
                int contentIndex = 0;
                foreach (var content in EnumerateCellContents(cell))
                {
                    if (TryGetBlockTableRecordIdFromContent(content) != ObjectId.Null)
                        blockContentIndexes.Add(contentIndex);

                    contentIndex++;
                }

                if (blockContentIndexes.Count > 0)
                {
                    // Try to set an attribute on the cell's block content.
                    foreach (var idx in blockContentIndexes)
                    {
                        foreach (var tag in CrossingBubbleAttributeTags)
                        {
                            if (TryInvokeSetBlockAttributeValue(table, row, col, idx, tag, crossingText))
                                return true;
                        }
                    }

                    // Fallback: try signature without content index (some AutoCAD versions).
                    foreach (var tag in CrossingBubbleAttributeTags)
                    {
                        if (TryInvokeSetBlockAttributeValue(table, row, col, null, tag, crossingText))
                            return true;
                    }

                    return false;
                }

                // Text-only cell.
                cell.TextString = crossingText;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryInvokeSetBlockAttributeValue(Table table, int row, int col, int? contentIndex, string tag, string value)
        {
            try
            {
                var t = table.GetType();

                if (contentIndex.HasValue)
                {
                    var mi5 = t.GetMethod(
                        "SetBlockAttributeValue",
                        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                        null,
                        new[] { typeof(int), typeof(int), typeof(int), typeof(string), typeof(string) },
                        null);

                    if (mi5 != null)
                    {
                        mi5.Invoke(table, new object[] { row, col, contentIndex.Value, tag, value });
                        return true;
                    }
                }

                var mi4 = t.GetMethod(
                    "SetBlockAttributeValue",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                    null,
                    new[] { typeof(int), typeof(int), typeof(string), typeof(string) },
                    null);

                if (mi4 != null)
                {
                    mi4.Invoke(table, new object[] { row, col, tag, value });
                    return true;
                }
            }
            catch
            {
                // ignore
            }

            return false;
        }

        // ===== Buttons (Designer wires to these; keep them) =====

        // Scan button: read-only refresh of the grid/duplicate UI (no writes back to DWG)
        private void btnRescan_Click(object sender, EventArgs e) => RescanRecords(applyToTables: false);

        private void btnApply_Click(object sender, EventArgs e) => ApplyChangesToDrawing();

        private void btnGenerateAllPages_Click(object sender, EventArgs e) => GenerateAllXingPages();

        private void btnGenerateAllLatLongTables_Click(object sender, EventArgs e) => GenerateWaterLatLongTables();

        private void btnGenerateOtherLatLongTables_Click(object sender, EventArgs e) => GenerateOtherLatLongTables();

        // DELETE SELECTED: does NOT renumber the remaining crossings.
        // - removes the block instances for the selected record
        // - removes matching rows from all recognized tables (Main/Page/LatLong)
        // - writes current grid back to blocks
        // - forces a real graphics refresh to avoid the Ã¢â‚¬Å“ghost rowÃ¢â‚¬Â
                
        private void btnDelete_Click(object sender, EventArgs e)
        {
            // Support deleting multiple selected rows.
            var rowIndices = GetSelectedRowIndices();

            var recordsToDelete = rowIndices
                .Select(i => gridCrossings.Rows[i].DataBoundItem as CrossingRecord)
                .Where(r => r != null)
                .Distinct()
                .ToList();

            // Fallback if selection came from CurrentRow only
            if (recordsToDelete.Count == 0)
            {
                var single = GetSelectedRecord();
                if (single != null)
                    recordsToDelete.Add(single);
            }

            if (recordsToDelete.Count == 0)
                return;

            string prompt;
            if (recordsToDelete.Count == 1)
            {
                prompt = $"Delete {recordsToDelete[0].Crossing}?";
            }
            else
            {
                var preview = string.Join(", ", recordsToDelete.Take(10).Select(r => r.Crossing));
                if (recordsToDelete.Count > 10)
                    preview += ", ...";

                prompt = $"Delete {recordsToDelete.Count} selected crossings?" + Environment.NewLine + Environment.NewLine + preview;
            }

            if (MessageBox.Show(prompt, "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            try
            {
                var idsToDelete = recordsToDelete
                    .SelectMany(r => r.AllInstances)
                    .Distinct()
                    .ToList();

                _repository.DeleteInstances(idsToDelete);

                var tableErrors = new List<string>();
                foreach (var rec in recordsToDelete)
                {
                    try
                    {
                        DeleteRowFromTables(rec.CrossingKey);
                    }
                    catch (Exception ex)
                    {
                        tableErrors.Add($"{rec.Crossing}: {ex.Message}");
                    }

                    _records.Remove(rec);
                }

                if (tableErrors.Count > 0)
                {
                    MessageBox.Show(
                        "Some table rows could not be deleted:" + Environment.NewLine + Environment.NewLine + string.Join(Environment.NewLine, tableErrors),
                        "Table delete warning",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }

                _isDirty = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Delete failed");
            }
        }



        private void BtnMoveUpOnClick(object sender, EventArgs e) => MoveSelectedCrossingRow(-1);

        private void BtnMoveDownOnClick(object sender, EventArgs e) => MoveSelectedCrossingRow(1);

        // Resequence X# values (X1..Xn) based on the current list order.
        // This does not move rows; it only renumbers them.
        private void BtnRenumberFromListOnClick(object sender, EventArgs e)
        {
            ResequenceCrossingsFromGridOrder();
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

                // A fresh scan becomes the new baseline for X# values.
                _pendingRenumberMap = null;

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
                        UpdateAllXingTablesFromGrid();               // safety XÃ¢â‚¬â€˜key pass (4/6Ã¢â‚¬â€˜col LAT/LONG)
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

        // --- REPLACE your existing ApplyChangesToDrawing() with this exact method ---
        // =======================
        // XingForm.cs  (inside class XingForm)
        // =======================

        // REPLACE your existing ApplyChangesToDrawing() with this full method.
        private void ApplyChangesToDrawing()
        {
            // Commit any in-grid edits first
            FlushPendingGridEdits();

            if (!ValidateRecords()) return;

            try
            {
                // Snapshot the grid BEFORE we touch the DWG; this is our Ã¢â‚¬Å“source of truthÃ¢â‚¬Â
                var snapshot = _records.ToList();

                // 1) Push grid -> blocks (your existing behavior)
                _repository.ApplyChanges(snapshot, _tableSync);                                             // updates block attributes etc.  (kept)  :contentReference[oaicite:4]{index=4}

                // 2) Push grid -> all RECOGNIZED tables (Main/Page/LatLong) via TableSync (kept)
                _tableSync.UpdateAllTables(_doc, snapshot);                                                 // kept  :contentReference[oaicite:5]{index=5}

                // 2b) If the user resequenced X#s (Move Up/Down, drag reorder, Renumber List Order),
                //     apply the pending X# map to ANY table bubble cells (layout variants / legacy tables).
                ApplyPendingRenumberMapToAnyTables();

                // 3) Safety pass on recognized tables keyed by X (kept)
                UpdateAllXingTablesFromGrid();                                                              // kept  :contentReference[oaicite:6]{index=6}

                // 4) If scanner found explicit LAT/LONG sources, update those directly (kept)
                _tableSync.UpdateLatLongSourceTables(_doc, snapshot);                                       // kept  :contentReference[oaicite:7]{index=7}

                // 5) NEW: robust LAT/LONG sweep that handles headerless/legacy tables.
                //    - uses TableSync.ResolveCrossingKey for X
                //    - finds LAT/LONG columns per *row* (labels or numeric ranges)
                //    - tags any touched table as LATLONG so future scans recognize it
                ReplaceLatLongInAnyTableRobust(snapshot);                                                   // NEW

                // 6) Visual refresh of any tables we just touched (kept)
                ForceRegenAllCrossingTablesInDwg();                                                         // kept  :contentReference[oaicite:8]{index=8}

                // 7) Re-read from DWG and allow the apply-to-drawing pass to push any resolver results.
                RescanRecords(applyToTables: true, suppressDuplicateUi: true);                             // kept  :contentReference[oaicite:9]{index=9}

                // 8) NEW: guard against Ã¢â‚¬Å“snap backÃ¢â‚¬Â only for LAT/LONG by reapplying the snapshotÃ¢â‚¬â„¢s LAT/LONG.
                //    This does not affect Owner/Desc/Location/DWG_REF (your original flow is preserved).
                OverrideLatLongFromSnapshotInGrid(snapshot);                                                // NEW

                _isDirty = false;
                MessageBox.Show("Crossing data applied to drawing.", "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // NEW: Update LAT/LONG in ANY table row that matches by X, even with no headings.
        //      - Reads X with TableSync.ResolveCrossingKey (robust).
        //      - Finds per-row LAT/LONG columns by label (Ã¢â‚¬Å“LATÃ¢â‚¬Â, Ã¢â‚¬Å“LONGÃ¢â‚¬Â, Ã¢â‚¬Å“LATITUDEÃ¢â‚¬Â, Ã¢â‚¬Å“LONGITUDEÃ¢â‚¬Â) or numeric range.
        //      - Writes only LAT and LONG to minimize side-effects.
        //      - Tags updated tables as LATLONG (so IdentifyTable will classify them next time).
        private void ReplaceLatLongInAnyTableRobust(IList<CrossingRecord> snapshot)
        {
            var doc = _doc ?? AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null || snapshot == null || snapshot.Count == 0) return;

            // Build a key -> record map (accept either Ã¢â‚¬Å“X12Ã¢â‚¬Â or bare digits, normalize both ways)
            var byX = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in snapshot)
            {
                if (r == null) continue;
                var k1 = TableSync.NormalizeKeyForLookup(r.Crossing);   // Ã¢â‚¬Å“X##Ã¢â‚¬Â or ""  (robust normalization)  :contentReference[oaicite:10]{index=10}
                var k2 = NormalizeXKey(r.Crossing);                     // fallback normalizer used elsewhere     :contentReference[oaicite:11]{index=11}
                if (!string.IsNullOrWhiteSpace(k1) && !byX.ContainsKey(k1)) byX[k1] = r;
                if (!string.IsNullOrWhiteSpace(k2) && !byX.ContainsKey(k2)) byX[k2] = r;
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
                        var table = tr.GetObject(entId, OpenMode.ForRead) as Table;
                        if (table == null) continue;

                        // We don't depend on IdentifyTable here: this pass is meant to catch legacy/headerless tables too.
                        // If IdentifyTable *does* work, great; if not, we still handle by row.
                        var cols = table.Columns.Count;
                        if (!(cols == 4 || cols >= 6)) continue;   // only LAT/LONG layouts
                        table.UpgradeOpen();

                        bool anyRowUpdated = false;
                        for (int row = 0; row < table.Rows.Count; row++)
                        {
                            // Robustly read X from Column 0 (text, mtext, or block attr)
                            string rawKey = string.Empty;
                            try { rawKey = TableSync.ResolveCrossingKey(table, row, 0); } catch { rawKey = string.Empty; }  // :contentReference[oaicite:12]{index=12}

                            var key = TableSync.NormalizeKeyForLookup(rawKey);
                            if (string.IsNullOrWhiteSpace(key))
                                key = NormalizeXKey(rawKey);

                            if (string.IsNullOrWhiteSpace(key) || !byX.TryGetValue(key, out var rec) || rec == null)
                                continue;

                            // Detect which columns on this row are LAT and LONG
                            if (!TryDetectLatLongColumns(table, row, out var latCol, out var longCol))
                                continue;

                            bool changed = false;
                            changed |= SetCellIfChanged(table, row, latCol, rec.Lat ?? string.Empty);
                            changed |= SetCellIfChanged(table, row, longCol, rec.Long ?? string.Empty);

                            if (changed) anyRowUpdated = true;
                        }

                        if (anyRowUpdated)
                        {
                            try { table.GenerateLayout(); } catch { }
                            NormalizeTableBorders(table);
                            try { table.RecordGraphicsModified(true); } catch { }

                            // Tag this table as LATLONG so future scans/updates classify it correctly
                            try
                            {
                                _tableFactory.TagTable(tr, table, TableSync.XingTableType.LatLong.ToString().ToUpperInvariant());  // :contentReference[oaicite:13]{index=13}
                            }
                            catch { /* tagging is best-effort */ }

                            // Force a redraw to avoid the Ã¢â‚¬Å“ghost rowÃ¢â‚¬Â look
                            ForceRegenTable(table);                                                                              // :contentReference[oaicite:14]{index=14}
                        }
                    }
                }

                tr.Commit();
            }

            // Non-undoable repaint
            try { doc.Editor?.Regen(); } catch { }
            try { AcadApp.UpdateScreen(); } catch { }
        }


        // NEW: After Rescan (which may ignore headerless LAT/LONG tables), put just the
        //      snapshotÃ¢â‚¬â„¢s LAT/LONG (and Zone if you want) back into the grid so the form
        //      shows the values you just applied. Everything else remains from the rescan.
        private void OverrideLatLongFromSnapshotInGrid(IList<CrossingRecord> snapshot)
        {
            if (snapshot == null || _records == null || _records.Count == 0) return;

            var snapByX = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in snapshot)
            {
                if (r == null) continue;
                var k = NormalizeXKey(r.Crossing);
                if (!string.IsNullOrWhiteSpace(k) && !snapByX.ContainsKey(k))
                    snapByX[k] = r;
            }

            foreach (var r in _records)
            {
                if (r == null) continue;
                var key = NormalizeXKey(r.Crossing);
                if (string.IsNullOrWhiteSpace(key)) continue;

                if (snapByX.TryGetValue(key, out var s) && s != null)
                {
                    if (!string.Equals(r.Lat, s.Lat, StringComparison.Ordinal)) r.Lat = s.Lat;
                    if (!string.Equals(r.Long, s.Long, StringComparison.Ordinal)) r.Long = s.Long;

                    // If you also want Zone to stick along with LAT/LONG, un-comment:
                    // if (!string.Equals(r.Zone, s.Zone, StringComparison.Ordinal)) r.Zone = s.Zone;
                }
            }

            _records.ResetBindings();
            gridCrossings.Refresh();
        }


        // NEW (local heuristic): Try to find the LAT and LONG columns on a data row.
        // Strategy (per row):
        //   1) If we see "LAT"/"LATITUDE" (or "LONG"/"LONGITUDE"), prefer the neighbor to the right (else left).
        //   2) Otherwise, pick numeric candidates by range: LAT in [-90..90], LONG in [-180..180].
        //   3) Prefer a LONG that lies to the right of the chosen LAT when both exist.
        private static bool TryDetectLatLongColumns(Table table, int row, out int latCol, out int longCol)
        {
            latCol = -1; longCol = -1;
            if (table == null || row < 0 || row >= table.Rows.Count) return false;

            int cols = table.Columns.Count;
            var latLabelCols = new List<int>();
            var longLabelCols = new List<int>();
            var latNumericCols = new List<int>();
            var longNumericCols = new List<int>();

            for (int c = 0; c < cols; c++)
            {
                string raw = string.Empty;
                try { raw = table.Cells[row, c]?.TextString ?? string.Empty; } catch { raw = string.Empty; }

                var txt = (raw ?? string.Empty).Trim();
                var up = txt.ToUpperInvariant();

                if (up.Contains("LATITUDE") || up == "LAT" || up.StartsWith("LAT"))
                    latLabelCols.Add(c);

                if (up.Contains("LONGITUDE") || up == "LONG" || up.StartsWith("LONG"))
                    longLabelCols.Add(c);

                if (double.TryParse(txt, NumberStyles.Float, CultureInfo.InvariantCulture, out double val))
                {
                    if (val >= -90.0 && val <= 90.0) latNumericCols.Add(c);
                    if (val >= -180.0 && val <= 180.0) longNumericCols.Add(c);
                }
            }

            // Label -> neighbor
            foreach (var c in latLabelCols)
            {
                if (c + 1 < cols) { latCol = c + 1; break; }
                if (c - 1 >= 0) { latCol = c - 1; break; }
            }
            foreach (var c in longLabelCols)
            {
                if (c + 1 < cols) { longCol = c + 1; break; }
                if (c - 1 >= 0) { longCol = c - 1; break; }
            }

            // Fill from numeric candidates if needed
            if (latCol < 0 && latNumericCols.Count > 0) latCol = latNumericCols[0];

            if (longCol < 0 && longNumericCols.Count > 0)
            {
                // Prefer a LONG to the right of LAT if we already picked a LAT
                int pick = -1;
                foreach (var c in longNumericCols)
                {
                    if (latCol < 0 || c > latCol) { pick = c; break; }
                }
                longCol = (pick >= 0) ? pick : longNumericCols[0];
            }

            // If both landed on the same column, split LONG to an alternative if possible
            if (latCol >= 0 && longCol == latCol && longNumericCols.Count > 1)
            {
                foreach (var c in longNumericCols)
                {
                    if (c != latCol) { longCol = c; break; }
                }
            }

            return latCol >= 0 && longCol >= 0 && latCol < cols && longCol < cols && latCol != longCol;
        }


        // --- NEW helper: only touches LAT, LONG (and ZONE/DWG_REF column in 6Ã¢â‚¬â€˜col LAT/LONG tables) ---
        private void BruteForceReplaceLatLongEverywhere(IList<CrossingRecord> snapshot)
        {
            var doc = _doc ?? AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null || snapshot == null || snapshot.Count == 0) return;

            // Build byÃ¢â‚¬â€˜X map from the current grid snapshot ("X3" etc.)
            var byX = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in snapshot)
            {
                var key = NormalizeXKey(r?.Crossing);
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
                        var table = tr.GetObject(entId, OpenMode.ForRead) as Table;
                        if (table == null) continue;

                        // Treat classic 4Ã¢â‚¬â€˜ or 6Ã¢â‚¬â€˜column layouts as LAT/LONG candidates (ID|DESC|LAT|LONG) or (ID|DESC|ZONE|LAT|LONG|DWG_REF)
                        var cols = table.Columns.Count;
                        if (cols != 4 && cols != 6) continue;

                        // Use the same startÃ¢â‚¬â€˜row logic you already rely on elsewhere
                        int startRow = 1;
                        try { startRow = TableSync.FindLatLongDataStartRow(table); } // robust header finder. :contentReference[oaicite:7]{index=7}
                        catch { startRow = 1; }
                        if (startRow <= 0) startRow = 1;

                        table.UpgradeOpen();

                        bool anyUpdated = false;
                        int zoneCol = (cols >= 6) ? 2 : -1;
                        int latCol = (cols >= 6) ? 3 : 2;
                        int lngCol = (cols >= 6) ? 4 : 3;
                        int dwgCol = (cols >= 6) ? 5 : -1;

                        for (int row = startRow; row < table.Rows.Count; row++)
                        {
                            // Column 0 key: attributeÃ¢â‚¬â€˜first, fall back to visible token
                            string xRaw = ReadXFromCellAttributeOnly(table, row, tr);
                            if (string.IsNullOrWhiteSpace(xRaw))
                            {
                                try
                                {
                                    var txt = table.Cells[row, 0]?.TextString ?? string.Empty;
                                    xRaw = ExtractXToken(txt);
                                }
                                catch { /* ignore */ }
                            }

                            var xKey = NormalizeXKey(xRaw);
                            if (string.IsNullOrWhiteSpace(xKey) || !byX.TryGetValue(xKey, out var rec) || rec == null)
                                continue;

                            // Write LAT/LONG (and Zone/DWG_REF if the 6Ã¢â‚¬â€˜col layout is present)
                            string zoneLabel = string.IsNullOrWhiteSpace(rec.Zone)
                                ? string.Empty
                                : string.Format(CultureInfo.InvariantCulture, "ZONE {0}", rec.Zone.Trim());

                            anyUpdated |= SetCellIfChanged(table, row, latCol, rec.Lat);
                            anyUpdated |= SetCellIfChanged(table, row, lngCol, rec.Long);
                            if (zoneCol >= 0) anyUpdated |= SetCellIfChanged(table, row, zoneCol, zoneLabel);
                            if (dwgCol >= 0 && !string.IsNullOrWhiteSpace(rec.DwgRef)) anyUpdated |= SetCellIfChanged(table, row, dwgCol, rec.DwgRef);
                        }

                        if (anyUpdated)
                        {
                            try { table.GenerateLayout(); } catch { }
                            NormalizeTableBorders(table);
                            try { table.RecordGraphicsModified(true); } catch { }
                        }
                    }
                }

                tr.Commit();
            }

            // NonÃ¢â‚¬â€˜undoable repaint paths
            try { doc.Editor?.Regen(); } catch { }
            try { AcadApp.UpdateScreen(); } catch { }
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
        /// Ends with a hard refresh to avoid the transient Ã¢â‚¬Å“ghost rowÃ¢â‚¬Â look.
        /// Update recognised crossing tables (MAIN/PAGE/LATLONG) from the current grid.
        /// Never modifies Column 0 (the bubble/X); only columns B..E are written.
        /// Ends with a hard refresh to avoid the transient Ã¢â‚¬Å“ghost rowÃ¢â‚¬Â look.
        /// Update recognised crossing tables (MAIN, PAGE, LATLONG) from the current grid.
        /// - Matches rows by X-key (attribute-first, with text fallback).
        /// - Never writes Column 0 (bubble) in any table.
        /// - For LAT/LONG, supports both 4Ã¢â‚¬â€˜col (ID, DESC, LAT, LONG) and 6Ã¢â‚¬â€˜col (ID, DESC, ZONE, LAT, LONG, DWG_REF).
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
                                CommandLogger.Log(ed, $"Row {row} -> NO MATCH (key='{xRaw}')");
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
                                // 4Ã¢â‚¬â€˜col: ID | DESC | LAT | LONG
                                // 6Ã¢â‚¬â€˜col: ID | DESC | ZONE | LAT | LONG | DWG_REF
                                var cols = table.Columns.Count;

                                // Description (col 1 is common)
                                changed |= SetCellIfChanged(table, row, 1, rec.Description);

                                // Zone label for 6Ã¢â‚¬â€˜col tables (e.g., "ZONE 11")
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
                            NormalizeTableBorders(table);
                            try { table.RecordGraphicsModified(true); } catch { }
                            ForceRegenTable(table);
                        }

                        var handleHex = table.ObjectId.Handle.Value.ToString("X");
                        CommandLogger.Log(ed, $"Table {handleHex}: {kind.ToString().ToUpperInvariant()} matched={matched} updated={updated}");
                    }
                }

                tr.Commit();
            }

            try { doc.Editor?.Regen(); } catch { }
            try { AcadApp.UpdateScreen(); } catch { }
        }

        // Set borders on a single cell via grid visibility
        private static MethodInfo _setGridVisibilityBool;
        private static MethodInfo _setGridVisibilityEnum;
        private static Type _visibilityEnumType;

        private static void SetCellBorders(Table t, int r, int c, bool top, bool right, bool bottom, bool left)
        {
            if (t == null) return;
            TrySetGridVisibility(t, r, c, GridLineType.HorizontalTop, top);
            TrySetGridVisibility(t, r, c, GridLineType.VerticalRight, right);
            TrySetGridVisibility(t, r, c, GridLineType.HorizontalBottom, bottom);
            TrySetGridVisibility(t, r, c, GridLineType.VerticalLeft, left);
        }

        private static void TrySetGridVisibility(Table t, int row, int col, GridLineType line, bool visible)
        {
            if (t == null) return;

            try
            {
                var tableType = t.GetType();
                if (_setGridVisibilityBool == null)
                {
                    _setGridVisibilityBool = tableType.GetMethod(
                        "SetGridVisibility",
                        new[] { typeof(int), typeof(int), typeof(GridLineType), typeof(bool) });
                }

                if (_setGridVisibilityBool != null)
                {
                    _setGridVisibilityBool.Invoke(t, new object[] { row, col, line, visible });
                    return;
                }

                if (_visibilityEnumType == null)
                {
                    _visibilityEnumType = tableType.Assembly.GetType("Autodesk.AutoCAD.DatabaseServices.Visibility");
                }
                if (_visibilityEnumType == null)
                    return;

                if (_setGridVisibilityEnum == null)
                {
                    _setGridVisibilityEnum = tableType.GetMethod(
                        "SetGridVisibility",
                        new[] { typeof(int), typeof(int), typeof(GridLineType), _visibilityEnumType });
                }

                if (_setGridVisibilityEnum == null)
                    return;

                object enumValue = null;
                try
                {
                    var field = _visibilityEnumType.GetField(visible ? "Visible" : "Invisible", BindingFlags.Public | BindingFlags.Static);
                    enumValue = field?.GetValue(null);
                }
                catch
                {
                    enumValue = null;
                }

                if (enumValue == null)
                {
                    try
                    {
                        enumValue = Enum.Parse(_visibilityEnumType, visible ? "Visible" : "Invisible");
                    }
                    catch
                    {
                        return;
                    }
                }

                _setGridVisibilityEnum.Invoke(t, new[] { (object)row, (object)col, (object)line, enumValue });
            }
            catch
            {
                // best effort: ignore visibility failures
            }
        }

        private static int TryFindDataStartRow(Table t)
        {
            try
            {
                // Prefer the existing helper if available
                int r = TableSync.FindLatLongDataStartRow(t);
                return (r < 0 ? 0 : r);
            }
            catch { return 0; }
        }

        // strip inline MTEXT formatting codes like \A1;, \H2.5x;, \C2;, {Ã¢â‚¬Â¦}, etc.
        private static readonly Regex InlineCode = new Regex(@"\\[A-Za-z][^;]*;|{[^}]*}");

        // These regexes remove MTEXT commands and special codes from cell strings.
        private static readonly Regex ResidualCode = new Regex(@"\\[^{}]", RegexOptions.Compiled);
        private static readonly Regex SpecialCode = new Regex("%%[^\\s]+", RegexOptions.Compiled);

        // Strip inline MTEXT formatting commands (\H, \A, \C, etc.) and braces.
        // This is the method that was missing and causing CS0103.
        private static string StripMTextFormatting(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            // InlineCode is already defined in your XingForm ( \[A-Za-z][^;]*;|{[^}]*} )
            var withoutCommands = InlineCode.Replace(value, string.Empty);
            var withoutResidual = ResidualCode.Replace(withoutCommands, string.Empty);
            var withoutSpecial = SpecialCode.Replace(withoutResidual, string.Empty);
            return withoutSpecial.Replace("{", string.Empty).Replace("}", string.Empty);
        }
        // Normalize all cells: heading rows = underline only; data rows = full box
        private static void NormalizeTableBorders(Table t)
        {
            int rows = t.Rows.Count;
            int cols = t.Columns.Count;
            int dataStart = TryFindDataStartRow(t);

            for (int r = 0; r < rows; r++)
            {
                bool sectionTitle = IsSectionTitleRow(t, r);
                bool columnHeader =
                    !sectionTitle &&
                    (
                        (dataStart > 0 && r == dataStart - 1) ||
                        (r > 0 && IsSectionTitleRow(t, r - 1))
                    );

                for (int c = 0; c < cols; c++)
                {
                    if (sectionTitle)
                    {
                        bool showTop = r > 0;
                        SetCellBorders(t, r, c, top: showTop, right: false, bottom: true, left: false);
                    }
                    else if (columnHeader)
                    {
                        SetCellBorders(t, r, c, top: false, right: false, bottom: true, left: false);
                    }
                    else
                    {
                        SetCellBorders(t, r, c, top: true, right: true, bottom: true, left: true);
                    }
                }
            }

            // Final pass: if the next row is a section title, make sure the current row
            // has a full-width bottom border so there is no gap.
            ApplySectionSeparators(t);
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
            var peo = new PromptEntityOptions("\nSelect a crossing table (Main/Page/Lat-Long):");
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
                    CommandLogger.Log(ed, "Selection was not a Table.");
                    return false;
                }

                var kind = _tableSync.IdentifyTable(table, tr);
                if (kind == TableSync.XingTableType.Unknown)
                {
                    CommandLogger.Log(ed, "Could not determine table type.");
                    return false;
                }

                if (kind == TableSync.XingTableType.LatLong)
                {
                    CommandLogger.Log(ed, "Lat/Long table detected; merging coordinates into the grid.");
                    gridChanged = MergeLatLongTableIntoGrid(table, ed);
                }
                else
                {
                    var byX = new Dictionary<string, (string Raw, string Owner, string Desc, string Loc, string Dwg, int Row)>(StringComparer.OrdinalIgnoreCase);

                    CommandLogger.Log(ed, "--- TABLE VALUES (parsed) ---");
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
                        CommandLogger.Log(ed, $"[T] row {r}: X='{xKey}' owner='{owner}' desc='{desc}' loc='{loc}' dwg='{dwg}'");

                        if (!byX.ContainsKey(xKey))
                            byX[xKey] = (xRaw?.Trim() ?? string.Empty, owner, desc, loc, dwg, r);
                        else
                            CommandLogger.Log(ed, $"[!] Duplicate X '{xKey}' in table (row {r}). First occurrence kept.");
                    }

                    CommandLogger.Log(ed, $"Table rows indexed by X = {byX.Count}");
                    CommandLogger.Log(ed, "--- TABLEÃ¢â€ â€™FORM UPDATES (X-only) ---");

                    int matched = 0, updated = 0, added = 0, noMatch = 0;
                    var matchedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var rec in _records)
                    {
                        var xKey = NormalizeXKey(rec.Crossing);
                        if (string.IsNullOrWhiteSpace(xKey) || !byX.TryGetValue(xKey, out var src))
                        {
                            CommandLogger.Log(ed, $"[!] {rec.Crossing}: no matching X in table.");
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
                        matchedKeys.Add(xKey);
                        if (changed)
                        {
                            updated++;
                            CommandLogger.Log(ed, $"[U] {rec.Crossing}: grid updated from table (row {src.Row}).");
                        }
                        else
                        {
                            CommandLogger.Log(ed, $"[=] {rec.Crossing}: no changes needed (already matches).");
                        }
                    }

                    foreach (var kvp in byX)
                    {
                        if (matchedKeys.Contains(kvp.Key))
                            continue;

                        var src = kvp.Value;
                        var crossingLabel = string.IsNullOrWhiteSpace(src.Raw) ? kvp.Key : src.Raw;
                        var newRecord = new CrossingRecord
                        {
                            Crossing = crossingLabel,
                            Owner = src.Owner,
                            Description = src.Desc,
                            Location = src.Loc,
                            DwgRef = src.Dwg
                        };

                        if (kind != TableSync.XingTableType.Main)
                        {
                            newRecord.Location = string.Empty;
                            newRecord.DwgRef = string.Empty;
                        }

                        _records.Add(newRecord);
                        added++;
                        CommandLogger.Log(ed, $"[+] {newRecord.Crossing}: added to grid from table (row {src.Row}).");
                    }

                    CommandLogger.Log(ed, $"Match Table -> grid only (X-only): matched={matched}, updated={updated}, added={added}, noMatch={noMatch}");

                    gridChanged = (updated > 0) || (added > 0);
                }

                tr.Commit();
            } // end read/lock

            _records.ResetBindings();
            gridCrossings.Refresh();
            _isDirty = true;

            // === Persist to DWG + tables if requested ===
            if (persistAfterMatch && gridChanged)
            {
                try
                {
                    // 1) Write block attributes from the grid
                    _repository.ApplyChanges(_records.ToList(), _tableSync);

                    // 2) Update MAIN/PAGE/LATLONG tables from the grid (B..E only Ã¢â‚¬â€ never col 0)
                    UpdateAllXingTablesFromGrid();

                    // 3) Reload grid from DWG (no table writes)
                    RescanRecords(applyToTables: false);

                    _isDirty = false;
                    CommandLogger.Log(ed, "MATCH TABLE: applied changes to DWG (blocks & tables).");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return gridChanged;
        }

        private bool MergeLatLongTableIntoGrid(Table table, Editor ed)
        {
            if (table == null) return false;

            // Use XingRepository to parse lat/long rows from the selected table.
            var rows = new List<XingRepository.LatLongRowInfo>();
            XingRepository.CollectLatLongRows(table, rows);

            // Build a lookup dictionary from CrossingRecord.CrossingKey to CrossingRecord.
            var recordMap = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            foreach (var record in _records)
            {
                if (record == null) continue;

                var key = TableSync.NormalizeKeyForLookup(record.Crossing);
                if (string.IsNullOrEmpty(key) || recordMap.ContainsKey(key))
                    continue;

                recordMap[key] = record;
            }

            // Apply the lat/long rows to the existing records (this updates Lat, Long, Zone, DwgRef).
            XingRepository.ApplyLatLongRows(rows, recordMap);

            // Add any new records for rows that didnÃ¢â‚¬â„¢t match existing crossings.
            foreach (var row in rows)
            {
                var key = TableSync.NormalizeKeyForLookup(row.Crossing);
                if (string.IsNullOrEmpty(key) || recordMap.ContainsKey(key))
                    continue;

                var newRecord = new CrossingRecord
                {
                    Crossing = string.IsNullOrWhiteSpace(row.Crossing) ? key : row.Crossing,
                    Description = row.Description ?? string.Empty,
                    DwgRef = row.DwgRef ?? string.Empty,
                    Lat = row.Latitude?.Trim() ?? string.Empty,
                    Long = row.Longitude?.Trim() ?? string.Empty,
                    Zone = row.Zone?.Trim() ?? string.Empty,
                    Owner = string.Empty,
                    Location = string.Empty
                };
                _records.Add(newRecord);
                recordMap[key] = newRecord;
                CommandLogger.Log(ed, $"[+] {newRecord.Crossing}: added from lat/long table.");
            }

            // Refresh the grid
            _records.ResetBindings();
            gridCrossings.Refresh();
            _isDirty = true;

            // Return true if any record was updated or added
            return rows.Count > 0;
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
                return s.Trim();
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

                        // Find the first *data* row to start from.
                        int startRow = 0;
                        if (kind == TableSync.XingTableType.LatLong)
                        {
                            try { startRow = TableSync.FindLatLongDataStartRow(table); }
                            catch { startRow = 0; }
                            if (startRow < 0) startRow = 0;
                        }
                        else
                        {
                            // MAIN/PAGE: if row 0 isnÃ¢â‚¬â„¢t an X row, advance until we hit a row that looks like data.
                            startRow = 0;
                            while (startRow < table.Rows.Count)
                            {
                                string probe = string.Empty;
                                try
                                {
                                    // attribute-first
                                    probe = ReadXFromCellAttributeOnly(table, startRow, tr);
                                    if (string.IsNullOrWhiteSpace(probe))
                                    {
                                        // text fallback
                                        var txt = table.Cells[startRow, 0]?.TextString ?? string.Empty;
                                        probe = ExtractXToken(txt);
                                    }
                                }
                                catch { probe = string.Empty; }

                                if (!string.IsNullOrWhiteSpace(NormalizeXKey(probe)))
                                    break; // first data row
                                startRow++;
                            }
                        }

                        for (int row = startRow; row < table.Rows.Count; row++)
                        {
                            string xRaw = ReadXFromCellAttributeOnly(table, row, tr);
                            if (string.IsNullOrWhiteSpace(xRaw))
                            {
                                try
                                {
                                    var txt = table.Cells[row, 0]?.TextString ?? string.Empty;
                                    xRaw = ExtractXToken(txt);
                                }
                                catch { xRaw = string.Empty; }
                            }

                            var key = NormalizeXKey(xRaw);
                            if (!normalized.Equals(key, StringComparison.OrdinalIgnoreCase))
                                continue;

                            bool deleted = false;

                            // 1) Preferred: hard delete the row
                            try
                            {
                                if (row >= 0 && row < table.Rows.Count)
                                {
                                    table.DeleteRows(row, 1);
                                    deleted = true;
                                }
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                            {
                                // eInvalidInput often means merged/title/header row; fall through to clear.
                                if (ex.Message.IndexOf("eInvalidInput", StringComparison.OrdinalIgnoreCase) < 0)
                                    throw; // different failure: rethrow
                            }
                            catch
                            {
                                // unknown; try clear fallback
                            }

                            // 2) Fallback: clear the row cells so we donÃ¢â‚¬â„¢t leave stale data behind
                            if (!deleted && row >= 0 && row < table.Rows.Count)
                            {
                                try
                                {
                                    int cols = table.Columns.Count;
                                    for (int c = 0; c < cols; c++)
                                    {
                                        // Never touch Column 0Ã¢â‚¬â„¢s block content; just blank the visible text.
                                        // (If Column 0 is a block cell, TextString set is harmless/no-op.)
                                        SetCellIfChanged(table, row, c, string.Empty);
                                    }
                                }
                                catch { /* best effort */ }
                            }

                            try { table.GenerateLayout(); } catch { }
                            NormalizeTableBorders(table);
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

        // Modify your btnDelete_Click handler so it doesnÃ¢â‚¬â„¢t renumber bubbles
        // and calls DeleteRowFromTables and UpdateAllXingTablesFromGrid.

        private CrossingRecord GetSelectedRecord()
        {
            if (gridCrossings.CurrentRow == null) return null;
            return gridCrossings.CurrentRow.DataBoundItem as CrossingRecord;
        }

        private List<int> GetSelectedRowIndices()
        {
            var set = new HashSet<int>();

            foreach (DataGridViewRow row in gridCrossings.SelectedRows)
            {
                if (row == null || row.IsNewRow)
                    continue;
                if (row.Index >= 0)
                    set.Add(row.Index);
            }

            foreach (DataGridViewCell cell in gridCrossings.SelectedCells)
            {
                if (cell == null)
                    continue;
                if (cell.RowIndex >= 0)
                    set.Add(cell.RowIndex);
            }

            if (gridCrossings.CurrentCell != null && gridCrossings.CurrentCell.RowIndex >= 0)
                set.Add(gridCrossings.CurrentCell.RowIndex);

            return set.OrderBy(i => i).ToList();
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

                // Mark graphics Ã¢â‚¬Å“dirtyÃ¢â‚¬Â so the display list is rebuilt
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
            if (_isAwaitingRenumber) return;

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("No active document.");
                return;
            }

            _isAwaitingRenumber = true;

            try
            {
                // This renumber runs directly (no external command), so any pending table-renumber map
                // from drag/move operations must NOT be applied afterwards.
                _pendingRenumberMap = null;

                var ed = doc.Editor;

                // 1) Select the RNC path polyline
                var peo = new PromptEntityOptions("\nSelect the RNC path polyline (Create RNC Path): ");
                peo.SetRejectMessage("\nPlease select a polyline.");
                peo.AddAllowedClass(typeof(Polyline), exactMatch: false);

                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK)
                    return;

                List<Point3d> vertices;
                ObjectId spaceId;

                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    var pl = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;
                    if (pl == null)
                    {
                        MessageBox.Show("Selected entity is not a polyline.");
                        return;
                    }

                    if (pl.Closed)
                    {
                        MessageBox.Show("The selected polyline is closed. Please select the OPEN RNC path polyline.");
                        return;
                    }

                    spaceId = pl.OwnerId;

                    vertices = new List<Point3d>(pl.NumberOfVertices);
                    for (int i = 0; i < pl.NumberOfVertices; i++)
                    {
                        var pt = pl.GetPoint3dAt(i);
                        if (vertices.Count == 0 || vertices[vertices.Count - 1].DistanceTo(pt) > 1e-6)
                            vertices.Add(pt);
                    }

                    tr.Commit();
                }

                if (vertices == null || vertices.Count == 0)
                {
                    MessageBox.Show("RNC path has no usable vertices.");
                    return;
                }

                // 2) Scan crossings so we can update ALL instances (copies) of any crossing that changes
                var scan = _repository.ScanCrossings();
                if (scan == null || scan.Records == null || scan.Records.Count == 0)
                {
                    MessageBox.Show("No crossing bubbles found.");
                    return;
                }

                // 3) Build candidate instance list (only in the same space as the RNC path)
                var candidates = new List<(ObjectId Id, Point3d Pos, CrossingRecord Rec)>();

                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    foreach (var rec in scan.Records)
                    {
                        foreach (var id in rec.AllInstances)
                        {
                            var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                            if (br == null) continue;
                            if (br.OwnerId != spaceId) continue;

                            candidates.Add((id, br.Position, rec));
                        }
                    }

                    tr.Commit();
                }

                if (candidates.Count == 0)
                {
                    MessageBox.Show("No crossing bubbles were found in the same space as the selected RNC path.");
                    return;
                }

                // 4) Match each vertex -> closest bubble (within tolerance)
                const double maxDist = 5.0; // drawing units
                var usedInstanceIds = new HashSet<ObjectId>();
                var targetCrossingByRecord = new Dictionary<CrossingRecord, string>();
                var missingVertices = new List<int>();

                for (int i = 0; i < vertices.Count; i++)
                {
                    var vtx = vertices[i];

                    double bestDist = double.MaxValue;
                    (ObjectId Id, Point3d Pos, CrossingRecord Rec) best = default;
                    bool found = false;

                    foreach (var cand in candidates)
                    {
                        if (usedInstanceIds.Contains(cand.Id))
                            continue;

                        double d = vtx.DistanceTo(cand.Pos);
                        if (d < bestDist)
                        {
                            bestDist = d;
                            best = cand;
                            found = true;
                        }
                    }

                    if (!found || bestDist > maxDist)
                    {
                        missingVertices.Add(i + 1);
                        continue;
                    }

                    usedInstanceIds.Add(best.Id);

                    var newX = "X" + (i + 1);
                    if (targetCrossingByRecord.ContainsKey(best.Rec))
                    {
                        MessageBox.Show(
                            "RNC cancelled. The selected path appears to hit the same crossing bubble more than once." + Environment.NewLine +
                            "(A single bubble was matched to multiple vertices.)",
                            "RNC",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }

                    targetCrossingByRecord[best.Rec] = newX;
                }

                if (missingVertices.Count > 0)
                {
                    // Build a helpful list with vertex numbers and coordinates
                    var lines = missingVertices
                        .Select(v => {
                            var pt = vertices[v - 1];
                            return $"Vertex {v}: ({pt.X:0.###}, {pt.Y:0.###})";
                        })
                        .ToList();

                    MessageBox.Show(
                        "RNC cancelled. No bubble was found at these vertex locations:" + Environment.NewLine + Environment.NewLine +
                        string.Join(Environment.NewLine, lines) + Environment.NewLine + Environment.NewLine +
                        "Make sure each RNC path vertex is snapped to a crossing bubble insertion point.",
                        "RNC",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // 5) Apply updates (this writes all copies for each affected crossing)
                foreach (var kvp in targetCrossingByRecord)
                    kvp.Key.Crossing = kvp.Value;

                _repository.ApplyChanges(targetCrossingByRecord.Keys.ToList(), _tableSync);

                // 6) Refresh the UI list (no table syncing here; user regenerates tables after RNC)
                RescanRecords(applyToTables: false);

                MessageBox.Show(
                    $"RNC complete. Renumbered {targetCrossingByRecord.Count} crossings from path vertices.",
                    "RNC",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "RNC failed");
            }
            finally
            {
                _isAwaitingRenumber = false;
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

        private static bool IsPlaceholderDwgRef(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return true;
            }

            var trimmed = value.Trim();
            return string.Equals(trimmed, "-", StringComparison.Ordinal) ||
                   string.Equals(trimmed, "Ã¢â‚¬â€", StringComparison.Ordinal) ||
                   string.Equals(trimmed, "Ã¢â‚¬â€œ", StringComparison.Ordinal);
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
                .Where(s => !IsPlaceholderDwgRef(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s, DwgRefComparer)
                .ToList();

            if (refs.Count == 0)
            {
                MessageBox.Show(
                    "No DWG_REF values available.",
                    "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Build perÃ¢â‚¬â€˜DWG_REF list of distinct LOCATIONs
            var locMap = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (var dr in refs)
            {
                var locs = _records
                    .Where(r => string.Equals(
                        r.DwgRef ?? string.Empty, dr, StringComparison.OrdinalIgnoreCase))
                    .Select(r => r.Location ?? string.Empty)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                locMap[dr] = locs;
            }

            var options = PromptForAllPagesOptions(refs, locMap);
            if (options == null || options.Count == 0) return;

            // Order options by earliest crossing and DwgRef
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
                        // Clone layout from template
                        string actualName;
                        var layoutId = _layoutUtils.CloneLayoutFromTemplate(
                            _doc,
                            TemplatePath,
                            GetTemplateLayoutNameForDwgRef(opt.DwgRef),
                            string.Format(CultureInfo.InvariantCulture, "XING #{0}", opt.DwgRef),
                            out actualName);

                        // Update heading text
                        _layoutUtils.UpdatePlanHeadingText(_doc.Database, layoutId, opt.IncludeAdjacent);

                        // Replace LOCATION text
                        var loc = opt.SelectedLocation ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(loc))
                        {
                            if (_layoutUtils.TryFormatMeridianLocation(loc, out var formatted))
                                loc = formatted;

                            _layoutUtils.ReplacePlaceholderText(_doc.Database, layoutId, loc);
                        }

                        // Compute center of layout for page table placement
                        Point3d center;
                        using (var tr = _doc.Database.TransactionManager.StartTransaction())
                        {
                            center = GetLayoutCenter(layoutId, tr);
                            tr.Commit();
                        }

                        var dataRowCount = _records.Count(r =>
                            string.Equals(r.DwgRef ?? string.Empty, opt.DwgRef, StringComparison.OrdinalIgnoreCase));

                        _tableSync.GetPageTableSize(dataRowCount, out double totalW, out double totalH);

                        var waterLatLongRecords = _records
                            .Where(r => string.Equals(r.DwgRef ?? string.Empty, opt.DwgRef, StringComparison.OrdinalIgnoreCase))
                            .Where(HasWaterLatLongData)
                            .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                            .ToList();

                        var otherLatLongCandidates = _records
                            .Where(r => string.Equals(r.DwgRef ?? string.Empty, opt.DwgRef, StringComparison.OrdinalIgnoreCase))
                            .Where(HasOtherLatLongData)
                            .ToList();

                        var otherSections = BuildOwnerLatLongSections(otherLatLongCandidates, out var otherRowCount);

                        // Sort each sectionÃ¢â‚¬â„¢s records numerically and sort sections by earliest crossing
                        foreach (var sec in otherSections)
                        {
                            sec.Records = sec.Records
                                .OrderBy(r => BuildCrossingSortKey(r.Crossing))
                                .ToList();
                        }

                        otherSections = otherSections
                            .OrderBy(sec => BuildCrossingSortKey(sec.Records.First().Crossing))
                            .ToList();

                        // Get sizes of water and other tables
                        double waterTotalW = 0.0, waterTotalH = 0.0;
                        if (waterLatLongRecords.Count > 0)
                        {
                            _tableSync.GetLatLongTableSize(waterLatLongRecords.Count, out waterTotalW, out waterTotalH);
                        }

                        double otherTotalW = 0.0, otherTotalH = 0.0;
                        if (otherRowCount > 0)
                        {
                            _tableSync.GetLatLongTableSize(otherRowCount, out otherTotalW, out otherTotalH);
                        }

                        var pageInsert = new Point3d(
                            center.X - totalW / 2.0,
                            center.Y - totalH / 2.0,
                            0.0);

                        var waterInsert = waterLatLongRecords.Count > 0
                            ? new Point3d(
                                center.X - waterTotalW / 2.0,
                                pageInsert.Y + totalH + TableVerticalGap,
                                0.0)
                            : Point3d.Origin;

                        var otherInsert = otherRowCount > 0
                            ? new Point3d(
                                center.X - otherTotalW / 2.0,
                                pageInsert.Y + totalH + TableVerticalGap +
                                (waterLatLongRecords.Count > 0 ? waterTotalH + TableVerticalGap : 0.0),
                                0.0)
                            : Point3d.Origin;

                        using (var tr = _doc.Database.TransactionManager.StartTransaction())
                        {
                            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                            var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

                            // Insert the main page table
                            _tableSync.CreateAndInsertPageTable(_doc.Database, tr, btr, pageInsert, opt.DwgRef, _records);

                            // Insert water table if any
                            if (waterLatLongRecords.Count > 0)
                            {
                                _tableSync.CreateAndInsertLatLongTable(
                                    _doc.Database,
                                    tr,
                                    btr,
                                    waterInsert,
                                    waterLatLongRecords);
                            }

                            // Insert each ownerÃ¢â‚¬â„¢s other table separately
                            if (otherSections.Count > 0)
                            {
                                Point3d nextInsert = otherInsert;
                                foreach (var sec in otherSections)
                                {
                                    _tableSync.GetLatLongTableSize(sec.Records.Count, out _, out double hSec);

                                    _tableSync.CreateAndInsertLatLongTable(
                                        _doc.Database,
                                        tr,
                                        btr,
                                        nextInsert,
                                        null,
                                        OtherLatLongTableTitle,
                                        new List<TableSync.LatLongSection> { sec },
                                        includeTitleRow: false);

                                    nextInsert = new Point3d(
                                        nextInsert.X,
                                        nextInsert.Y - (hSec + TableVerticalGap),
                                        nextInsert.Z);
                                }
                            }

                            tr.Commit();
                        }

                        // Switch to the new layout (optional)
                        _layoutUtils.SwitchToLayout(_doc, actualName);
                    }

                    // Reorder layouts (best effort)
                    try
                    {
                        ReorderXingLayouts();
                    }
                    catch
                    {
                        // ignore errors
                    }
                }
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show($"Template not found: {TemplatePath}",
                    "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Reorders all layouts whose name begins with "XING #"
        /// by the numeric part of their name.  AutoCAD assigns new layouts
        /// arbitrary tab orders when they are created; this method ensures
        /// that the layout tabs appear in ascending numeric order
        /// (e.g., XING #1, XING #2, XING #3).  The Ã¢â‚¬Å“ModelÃ¢â‚¬Â tab (order 0)
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

                    var hasHydroToken =
                        description.IndexOf(HydroProfileDescriptionToken, StringComparison.OrdinalIgnoreCase) >= 0;
                    var owner = record?.Owner;
                    var hasHydroOwner = hasHydroToken && !string.IsNullOrWhiteSpace(owner) && HydroOwnerKeywords.Any(
                        keyword => owner.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
                    var hasHydroKeyword = HydroKeywords.Any(keyword =>
                        description.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

                    if (hasHydroOwner || hasHydroKeyword)
                    {
                        return HydroTemplateLayoutName;
                    }
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

        private void GenerateWaterLatLongTables()
        {
            var latRecords = _records
                .Where(HasWaterLatLongData)
                .Where(r => !string.IsNullOrWhiteSpace(r.DwgRef))
                .OrderBy(r => r.DwgRef ?? string.Empty, DwgRefComparer)
                .ThenBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();

            if (latRecords.Count == 0)
            {
                MessageBox.Show("No WATER LAT/LONG data available to create tables.", "Crossing Manager",
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

            var pointRes = editor.GetPoint("\nSpecify insertion point for WATER LAT/LONG table:");
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

        // Create separate LAT/LONG tables for each owner and stack them vertically.
        // Create separate LAT/LONG tables for each owner and stack them in ascending X order.
        // Create separate LAT/LONG tables for each owner, sorted by crossing number, and stack them vertically.
        private void GenerateOtherLatLongTables()
        {
            var eligibleRecords = _records
                .Where(HasOtherLatLongData)
                .Where(r => !string.IsNullOrWhiteSpace(r.DwgRef))
                .ToList();

            var sections = BuildOwnerLatLongSections(eligibleRecords, out var totalRowCount);

            if (totalRowCount == 0)
            {
                MessageBox.Show(
                    "No OTHER LAT/LONG data available to create tables.",
                    "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Sort each sectionÃ¢â‚¬â„¢s records by crossing number (ascending)
            foreach (var sec in sections)
            {
                sec.Records = sec.Records
                    .OrderBy(r => BuildCrossingSortKey(r.Crossing))
                    .ToList();
            }

            // Sort sections by the earliest crossing key in each section
            sections = sections
                .OrderBy(sec => BuildCrossingSortKey(
                    sec.Records.First().Crossing))
                .ToList();

            var editor = _doc.Editor;
            if (editor == null)
            {
                MessageBox.Show(
                    "Unable to access the AutoCAD editor.",
                    "Crossing Manager",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var pointRes = editor.GetPoint("\nSpecify insertion point for OTHER LAT/LONG table:");
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
                        var layoutDict = (DBDictionary)tr.GetObject(
                            _doc.Database.LayoutDictionaryId, OpenMode.ForRead);
                        if (layoutDict.Contains(layoutManager.CurrentLayout))
                        {
                            var layoutId = layoutDict.GetAt(layoutManager.CurrentLayout);
                            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                            btr = (BlockTableRecord)tr.GetObject(
                                layout.BlockTableRecordId, OpenMode.ForWrite);
                        }
                    }
                }
                catch
                {
                    btr = null;
                }

                if (btr == null)
                {
                    btr = (BlockTableRecord)tr.GetObject(
                        _doc.Database.CurrentSpaceId, OpenMode.ForWrite);
                }

                Point3d insertPoint = pointRes.Value;
                const double SectionGap = 10.0; // vertical gap between tables

                foreach (var section in sections)
                {
                    _tableSync.GetLatLongTableSize(section.Records.Count, out _, out double height);

                    _tableSync.CreateAndInsertLatLongTable(
                        _doc.Database,
                        tr,
                        btr,
                        insertPoint,
                        null,
                        OtherLatLongTableTitle,
                        new List<TableSync.LatLongSection> { section },
                        includeTitleRow: false);

                    insertPoint = new Point3d(
                        insertPoint.X,
                        insertPoint.Y - height - SectionGap,
                        insertPoint.Z);
                }

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
                .Where(s => !IsPlaceholderDwgRef(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s, DwgRefComparer)
                .ToList();

            if (!choices.Any())
            {
                MessageBox.Show(
                    "No DWG_REF values available.",
                    "Crossing Manager",
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

                // Clone layout and set up heading/location
                using (_doc.LockDocument())
                {
                    layoutId = _layoutUtils.CloneLayoutFromTemplate(
                        _doc,
                        TemplatePath,
                        GetTemplateLayoutNameForDwgRef(options.DwgRef),
                        string.Format(
                            CultureInfo.InvariantCulture,
                            "XING #{0}",
                            options.DwgRef),
                        out actualName);

                    _layoutUtils.UpdatePlanHeadingText(_doc.Database, layoutId, options.IncludeAdjacent);

                    var locationText = BuildLocationText(options.DwgRef);
                    if (!string.IsNullOrEmpty(locationText))
                        _layoutUtils.ReplacePlaceholderText(_doc.Database, layoutId, locationText);

                    _layoutUtils.SwitchToLayout(_doc, actualName);
                }

                // Collect water lat/long records
                var waterLatLongRecords = _records
                    .Where(r => string.Equals(r.DwgRef ?? string.Empty, options.DwgRef, StringComparison.OrdinalIgnoreCase))
                    .Where(HasWaterLatLongData)
                    .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                    .ToList();

                // Collect other lat/long records
                var otherLatLongCandidates = _records
                    .Where(r => string.Equals(r.DwgRef ?? string.Empty, options.DwgRef, StringComparison.OrdinalIgnoreCase))
                    .Where(HasOtherLatLongData)
                    .ToList();

                var otherSections = BuildOwnerLatLongSections(otherLatLongCandidates, out var otherRowCount);

                // Sort each sectionÃ¢â‚¬â„¢s records by crossing number
                foreach (var sec in otherSections)
                {
                    sec.Records = sec.Records
                        .OrderBy(r => BuildCrossingSortKey(r.Crossing))
                        .ToList();
                }

                // Sort sections by their earliest crossing key
                otherSections = otherSections
                    .OrderBy(sec => BuildCrossingSortKey(sec.Records.First().Crossing))
                    .ToList();

                // Insert PAGE table
                var pagePrompt = _doc.Editor.GetPoint(
                    "\nSpecify insertion point for Crossing Page Table:");
                if (pagePrompt.Status != PromptStatus.OK) return;

                using (_doc.LockDocument())
                using (var tr = _doc.Database.TransactionManager.StartTransaction())
                {
                    var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

                    _tableSync.CreateAndInsertPageTable(
                        _doc.Database,
                        tr,
                        btr,
                        pagePrompt.Value,
                        options.DwgRef,
                        _records);

                    tr.Commit();
                }

                // Insert WATER table (if any)
                if (waterLatLongRecords.Count > 0)
                {
                    var latPrompt = _doc.Editor.GetPoint(
                        "\nSpecify insertion point for WATER LAT/LONG table:");
                    if (latPrompt.Status == PromptStatus.OK)
                    {
                        using (_doc.LockDocument())
                        using (var tr = _doc.Database.TransactionManager.StartTransaction())
                        {
                            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                            var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

                            _tableSync.CreateAndInsertLatLongTable(
                                _doc.Database,
                                tr,
                                btr,
                                latPrompt.Value,
                                waterLatLongRecords);

                            tr.Commit();
                        }
                    }
                }

                // Insert separate OTHER tables (if any)
                if (otherSections.Count > 0)
                {
                    var otherPrompt = _doc.Editor.GetPoint(
                        "\nSpecify insertion point for OTHER LAT/LONG table:");
                    if (otherPrompt.Status == PromptStatus.OK)
                    {
                        using (_doc.LockDocument())
                        using (var tr = _doc.Database.TransactionManager.StartTransaction())
                        {
                            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                            var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

                            const double SectionGap = 10.0;
                            Point3d nextInsert = otherPrompt.Value;

                            foreach (var sec in otherSections)
                            {
                                _tableSync.GetLatLongTableSize(sec.Records.Count, out _, out double h);

                                _tableSync.CreateAndInsertLatLongTable(
                                    _doc.Database,
                                    tr,
                                    btr,
                                    nextInsert,
                                    null,
                                    OtherLatLongTableTitle,
                                    new List<TableSync.LatLongSection> { sec },
                                    includeTitleRow: false);

                                // move down by table height + gap
                                nextInsert = new Point3d(
                                    nextInsert.X,
                                    nextInsert.Y - (h + SectionGap),
                                    nextInsert.Z);
                            }

                            tr.Commit();
                        }
                    }
                }
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show($"Template not found: {TemplatePath}",
                    "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private static bool HasWaterLatLongData(CrossingRecord record)
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

        private static bool HasOtherLatLongData(CrossingRecord record)
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

            return TryMatchOwnerKeyword(record.Owner, out _);
        }

        private static bool TryMatchOwnerKeyword(string owner, out string keyword)
        {
            keyword = null;
            if (string.IsNullOrWhiteSpace(owner))
            {
                return false;
            }

            foreach (var candidate in HydroOwnerKeywords)
            {
                if (owner.IndexOf(candidate, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    keyword = candidate;
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Determines the heading/title for a single LAT/LONG table created from the currently selected record.
        /// <summary>
        /// Determines the heading/title for a single LAT/LONG table created from the currently selected record.
        ///
        /// Rules:
        /// - Known owners (NOVA/PGI/etc): "{OWNER} CROSSING INFORMATION"
        /// - Water crossings (by description/keywords): "WATER CROSSING INFORMATION"
        /// - Anything else / unknown: no title row
        ///
        /// NOTE: We intentionally do NOT treat a blank/"-" Owner as water by default, because many older drawings
        /// may have missing Owner data for non-water crossings. Instead we look for water/hydro keywords.
        /// </summary>
        private static void ResolveLatLongTitleForRecord(CrossingRecord record, out string title, out bool includeTitleRow, out bool includeColumnHeaders)
        {
            title = null;
            includeTitleRow = false;
            includeColumnHeaders = true;

            if (record == null)
            {
                includeColumnHeaders = false;
                return;
            }

            var owner = (record.Owner ?? string.Empty).Trim();
            var desc = record.Description ?? string.Empty;

            // 1) Known owner tables first (NOVA/PGI/etc)
            if (TryMatchOwnerKeyword(owner, out var keyword))
            {
                title = string.Format(CultureInfo.InvariantCulture, "{0} CROSSING INFORMATION", keyword.ToUpperInvariant());
                includeTitleRow = true;
                return;
            }

            // 2) Water/hydro by description keywords
            var looksLikeWater = false;

            if (!string.IsNullOrWhiteSpace(desc))
            {
                // Hydro keywords (Watercourse/Creek/River) plus a generic 'water' fallback.
                if (desc.IndexOf("water", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    looksLikeWater = true;
                }
                else
                {
                    foreach (var k in HydroKeywords)
                    {
                        if (desc.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            looksLikeWater = true;
                            break;
                        }
                    }
                }
            }

            // Some users/projects store WATER directly in Owner.
            if (!looksLikeWater && owner.IndexOf("water", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                looksLikeWater = true;
            }

            if (looksLikeWater)
            {
                title = "WATER CROSSING INFORMATION";
                includeTitleRow = true;
                return;
            }

            // 3) Unknown type => no heading and no column headers.
            title = null;
            includeTitleRow = false;
            includeColumnHeaders = false;
        }


        private static IList<TableSync.LatLongSection> BuildOwnerLatLongSections(
            IEnumerable<CrossingRecord> records,
            out int totalRowCount)
        {
            totalRowCount = 0;
            var sections = new List<TableSync.LatLongSection>();

            if (records == null)
            {
                return sections;
            }

            foreach (var keyword in HydroOwnerKeywords)
            {
                var matching = records
                    .Where(r => r != null && TryMatchOwnerKeyword(r.Owner, out var match) &&
                                string.Equals(match, keyword, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(r => r.DwgRef ?? string.Empty, DwgRefComparer)
                    .ThenBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                    .ToList();

                if (matching.Count == 0)
                {
                    continue;
                }

                var header = string.Format(
                    CultureInfo.InvariantCulture,
                    "{0} CROSSING INFORMATION",
                    keyword.ToUpperInvariant());

                sections.Add(new TableSync.LatLongSection
                {
                    Header = header,
                    Records = matching
                });

                totalRowCount += 1 + matching.Count;
            }

            return sections;
        }
        /// <summary>
        /// Returns true if the specified row is a section title (e.g. "NOVA CROSSING INFORMATION"),
        /// even when the cell contains MTEXT formatting codes.
        /// </summary>
        private static bool IsSectionTitleRow(Table t, int r)
        {
            if (t == null || r < 0 || r >= t.Rows.Count)
                return false;

            try
            {
                // Read the raw text from column 0
                var raw = t.Cells[r, 0]?.TextString ?? string.Empty;
                // Remove MTEXT formatting commands and braces
                var plain = StripMTextFormatting(raw).Trim();
                // Check for the key phrase "CROSSING INFORMATION" (caseÃ¢â‚¬â€˜insensitive)
                return plain.IndexOf("CROSSING INFORMATION", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
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

            // Always create a new LAT/LONG table for the currently selected row.
            // (Older behavior users relied on.)
            var pointRes = _doc.Editor.GetPoint("\nSpecify insertion point for LAT/LONG table:");
            if (pointRes.Status != PromptStatus.OK) return;

            using (_doc.LockDocument())
            using (var tr = _doc.Database.TransactionManager.StartTransaction())
            {
                // Insert into the current layout's paper space (or model space if that's what
                // the current layout is).
                var layoutManager = LayoutManager.Current;
                var layoutDict = (DBDictionary)tr.GetObject(_doc.Database.LayoutDictionaryId, OpenMode.ForRead);
                var layoutId = layoutDict.GetAt(layoutManager.CurrentLayout);
                var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

                // Title row should reflect the crossing "type" (water vs known owner), not always WATER.
                ResolveLatLongTitleForRecord(record, out var titleOverride, out var includeTitleRow, out var includeColumnHeaders);

                _tableSync.CreateAndInsertLatLongTable(
                    _doc.Database,
                    tr,
                    btr,
                    pointRes.Value,
                    new List<CrossingRecord> { record },
                    titleOverride: titleOverride,
                    sections: null,
                    includeTitleRow: includeTitleRow,
                    includeColumnHeaders: includeColumnHeaders);
                tr.Commit();
            }

            try { _doc.Editor.Regen(); } catch { }
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
        /// <summary>
        /// After NormalizeTableBorders runs, ensure that every data row (identified by its first cell starting with "X")
        /// has a bottom border across all columns.
        /// This provides a separator below the last row of each owner section.
        /// </summary>
        private static void ApplySectionSeparators(Table t)
        {
            if (t == null) return;
            int rowCount = t.Rows.Count;
            int colCount = t.Columns.Count;

            for (int row = 0; row < rowCount - 1; row++)
            {
                try
                {
                    // Get the text from column 0, strip MTEXT formatting and whitespace
                    var raw = t.Cells[row, 0]?.TextString ?? string.Empty;
                    var text = StripMTextFormatting(raw).Trim();

                    // If the row starts with "X" (e.g. X1, X46), treat it as a data row and draw a bottom border
                    if (!string.IsNullOrEmpty(text) && text.StartsWith("X", StringComparison.OrdinalIgnoreCase))
                    {
                        for (int col = 0; col < colCount; col++)
                        {
                            TrySetGridVisibility(t, row, col, GridLineType.HorizontalBottom, true);
                        }
                    }
                }
                catch
                {
                    // Ignore rows that can't be inspected; best-effort
                }
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
                CommandLogger.Log(editor, "** Invalid zone Ã¢â‚¬â€œ enter 11 or 12. **", alsoToCommandBar: true);
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

/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////////
