using Autodesk.AutoCAD.ApplicationServices;  // for Document, CommandEventArgs
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using XingManager.Models;
using XingManager.Services;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using WinFormsFlowDirection = System.Windows.Forms.FlowDirection;
using System.Text.RegularExpressions;


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
        private BindingList<CrossingRecord> _records = new BindingList<CrossingRecord>();
        private IDictionary<ObjectId, DuplicateResolver.InstanceContext> _contexts = new Dictionary<ObjectId, DuplicateResolver.InstanceContext>();
        private bool _isDirty;
        private bool _isScanning;

        private const string TemplatePath = @"M:\\Drafting\\_CURRENT TEMPLATES\\Compass_Main.dwt";
        private const string TemplateLayoutName = "X";

        public XingForm(Document doc, XingRepository repository, TableSync tableSync, LayoutUtils layoutUtils, TableFactory tableFactory, Serde serde, DuplicateResolver duplicateResolver)
        {
            if (doc == null) throw new ArgumentNullException("doc");
            if (repository == null) throw new ArgumentNullException("repository");
            if (tableSync == null) throw new ArgumentNullException("tableSync");
            if (layoutUtils == null) throw new ArgumentNullException("layoutUtils");
            if (tableFactory == null) throw new ArgumentNullException("tableFactory");
            if (serde == null) throw new ArgumentNullException("serde");
            if (duplicateResolver == null) throw new ArgumentNullException("duplicateResolver");

            InitializeComponent();
            _doc = doc;
            _repository = repository;
            _tableSync = tableSync;
            _layoutUtils = layoutUtils;
            _tableFactory = tableFactory;
            _serde = serde;
            _duplicateResolver = duplicateResolver;

            ConfigureGrid();
        }

        public void LoadData()
        {
            RescanRecords();
        }

        public void RescanData()
        {
            RescanRecords();
        }

        public void ApplyToDrawing()
        {
            ApplyChangesToDrawing();
        }

        public void GenerateXingPageFromCommand()
        {
            GenerateXingPage();
        }

        public void CreateLatLongRowFromCommand()
        {
            CreateOrUpdateLatLongTable();
        }

        public void MatchTableFromCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            // Subscribe once; will auto-unsubscribe inside the handler
            doc.CommandEnded += OnCommandMatchTableDone;
            doc.CommandCancelled += OnCommandMatchTableDone;
            try
            {
                // In some versions you also have CommandFailed; subscribe if present in your SDK
                doc.CommandFailed += OnCommandMatchTableDone;
            }
            catch
            {
                // ignore if not available
            }

            // Run the command (space to execute)
            doc.SendStringToExecute("XING_MATCH_TABLE ", true, false, true);
        }

        private void OnCommandMatchTableDone(object sender, CommandEventArgs e)
        {
            try
            {
                if (!string.Equals(e.GlobalCommandName, "XING_MATCH_TABLE", StringComparison.OrdinalIgnoreCase))
                    return;

                var doc = sender as Document;
                if (doc != null)
                {
                    // Unhook to avoid duplicate subscriptions
                    doc.CommandEnded -= OnCommandMatchTableDone;
                    doc.CommandCancelled -= OnCommandMatchTableDone;
                    try
                    {
                        doc.CommandFailed -= OnCommandMatchTableDone;
                    }
                    catch
                    {
                    }
                }

                // Refresh the grid from the DWG **without** writing back to tables
                RescanRecords(applyToTables: false);
            }
            catch
            {
                // Best-effort; avoid surfacing handler errors
            }
        }

        public void RenumberSequentiallyFromCommand()
        {
            RenumberSequential();
            _isDirty = true;
        }

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
            {
                gridCrossings.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void GridCrossingsOnCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_isScanning) return;
            _isDirty = true;
        }

        private void btnRescan_Click(object sender, EventArgs e)
        {
            RescanRecords();
        }

        // ===== Rescan refreshes the grid, resolves dups, APPLIES to DWG, then re-reads from DWG =====
        // ===== Rescan resolves duplicates, APPLIES to DWG, then re-reads from DWG =====
        private void RescanRecords(bool applyToTables = true)
        {
            _isScanning = true;
            try
            {
                // 1) Read from DWG
                var result = _repository.ScanCrossings();
                _records = new BindingList<CrossingRecord>(result.Records.ToList());
                _contexts = result.InstanceContexts ?? new Dictionary<ObjectId, DuplicateResolver.InstanceContext>();
                gridCrossings.DataSource = _records;

                // 2) Resolve duplicates
                var ok = _duplicateResolver.ResolveDuplicates(_records, _contexts);
                if (!ok) { gridCrossings.Refresh(); _isDirty = false; return; }

                // 3) Persist only if requested
                if (applyToTables)
                {
                    try { _repository.ApplyChanges(_records.ToList(), _tableSync); }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    // 4) Re-read from DWG for fresh state
                    var post = _repository.ScanCrossings();
                    _records = new BindingList<CrossingRecord>(post.Records.ToList());
                    _contexts = post.InstanceContexts ?? new Dictionary<ObjectId, DuplicateResolver.InstanceContext>();
                    gridCrossings.DataSource = _records;
                }

                gridCrossings.Refresh();
                _isDirty = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { _isScanning = false; }
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            ApplyChangesToDrawing();
        }

        private void ApplyChangesToDrawing()
        {
            if (!ValidateRecords()) return;

            try
            {
                _repository.ApplyChanges(_records.ToList(), _tableSync);
                _isDirty = false;
                MessageBox.Show("Crossing data applied to drawing.", "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool ValidateRecords()
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var record in _records)
            {
                if (string.IsNullOrWhiteSpace(record.Crossing))
                {
                    MessageBox.Show("Each record must have a CROSSING value.", "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                var key = record.Crossing.Trim().ToUpperInvariant();
                if (!seen.Add(key))
                {
                    MessageBox.Show(string.Format(CultureInfo.InvariantCulture, "Duplicate CROSSING value '{0}' detected.", record.Crossing), "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (!ValidateLatLongValue(record.Lat) || !ValidateLatLongValue(record.Long))
                {
                    MessageBox.Show(string.Format(CultureInfo.InvariantCulture, "LAT/LONG values for {0} must be decimal numbers.", record.Crossing), "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        private void btnAdd_Click(object sender, EventArgs e)
        {
            var record = new CrossingRecord { Crossing = GenerateNextCrossingName() };
            _records.Add(record);
            PromptPlacement(record);
            _isDirty = true;
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            var opts = new PromptIntegerOptions("\nInsert crossing at position:")
            {
                AllowZero = false,
                AllowNegative = false,
                LowerLimit = 1,
                UpperLimit = _records.Count + 1,
                DefaultValue = _records.Count + 1
            };

            var res = _doc.Editor.GetInteger(opts);
            if (res.Status != PromptStatus.OK) return;

            var index = Math.Min(Math.Max(res.Value - 1, 0), _records.Count);
            ShiftCrossings(index, 1);

            var record = new CrossingRecord { Crossing = GenerateCrossingName(index) };
            _records.Insert(index, record);
            PromptPlacement(record);
            _isDirty = true;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            var record = GetSelectedRecord();
            if (record == null)
            {
                MessageBox.Show("Select a crossing to delete.", "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var confirm = MessageBox.Show(string.Format(CultureInfo.InvariantCulture, "Delete crossing {0}?", record.Crossing), "Crossing Manager", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            try
            {
                _repository.DeleteInstances(record.AllInstances);
                var index = _records.IndexOf(record);
                _records.Remove(record);
                ShiftCrossings(index, -1);
                _isDirty = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnRenumber_Click(object sender, EventArgs e)
        {
            RenumberSequential();
            _isDirty = true;
        }

        // Button: MATCH TABLE  ->  Grid-only merge from selected table (no DWG/table writes)
        private void btnMatchTable_Click(object sender, EventArgs e)
        {
            MatchTableIntoGrid();   // <- updates the grid only
        }

        private void btnGeneratePage_Click(object sender, EventArgs e)
        {
            GenerateXingPage();
        }

        private void btnLatLong_Click(object sender, EventArgs e)
        {
            CreateOrUpdateLatLongTable();
        }

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
                        MessageBox.Show("Export complete.", "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void MergeImportedRecords(IEnumerable<CrossingRecord> imported)
        {
            var map = _records.ToDictionary(r => r.CrossingKey, r => r, StringComparer.OrdinalIgnoreCase);
            foreach (var record in imported)
            {
                CrossingRecord existing;
                if (map.TryGetValue(record.CrossingKey, out existing))
                {
                    existing.Owner = record.Owner;
                    existing.Description = record.Description;
                    existing.Location = record.Location;
                    existing.DwgRef = record.DwgRef;
                    existing.Lat = record.Lat;
                    existing.Long = record.Long;
                }
                else
                {
                    _records.Add(record);
                }
            }
        }

        private void PromptPlacement(CrossingRecord record)
        {
            var point = _doc.Editor.GetPoint("\nSpecify insertion point for new crossing:");
            if (point.Status != PromptStatus.OK)
            {
                MessageBox.Show("Crossing created in the list only. Apply will skip until placed.", "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                var id = _repository.InsertCrossing(record, point.Value);
                record.AllInstances.Add(id);
                record.CanonicalInstance = id;
                _contexts[id] = new DuplicateResolver.InstanceContext
                {
                    ObjectId = id,
                    SpaceName = "Model",
                    Owner = record.Owner,
                    Description = record.Description,
                    Location = record.Location,
                    DwgRef = record.DwgRef,
                    Lat = record.Lat,
                    Long = record.Long
                };
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private CrossingRecord GetSelectedRecord()
        {
            if (gridCrossings.CurrentRow == null) return null;
            return gridCrossings.CurrentRow.DataBoundItem as CrossingRecord;
        }

        private string GenerateNextCrossingName()
        {
            var max = 0;
            foreach (var record in _records)
            {
                var token = CrossingRecord.ParseCrossingNumber(record.Crossing);
                if (token.Number > max) max = token.Number;
            }
            return string.Format(CultureInfo.InvariantCulture, "X{0}", max + 1);
        }
        // ===== GRID-ONLY: read a selected table and merge its values into the form =====
        // Never writes to DWG or to the table. Updates only the in-memory BindingList.
        // ===== GRID‑ONLY: read a selected table and merge values into the form =====
        // This version:
        //   * reads Column A strictly from a BLOCK ATTRIBUTE only,
        //   * reads B..E as plain text,
        //   * merges by X# only (no heuristics), and
        //   * never writes to DWG or to the table — grid only.
        // ===== GRID‑ONLY: read a selected table and merge values into the form =====
        // Reads Column A strictly from a BLOCK ATTRIBUTE, B..E as plain text,
        // merges by X only (no heuristics), never writes to DWG/tables.
        // ===== GRID‑ONLY: read a selected table and merge values into the form =====
        // Reads Column A from the block in the cell (attribute first, then block def text/name),
        // reads B..E as plain text, and merges STRICTLY BY X only. No DWG/table writes.
        // ===== GRID‑ONLY: read a selected table and merge values into the form =====
        // Reads Column A from the BLOCK attribute inside the cell (strict), B..E as plain text,
        // merges by X only (no heuristics). Never writes to DWG/tables.
        // ===== GRID‑ONLY: read a selected table and merge values into the form =====
        // Reads Column A from the block in the cell (attribute on the cell *content*),
        // reads B..E as plain text, and merges STRICTLY BY X only. Never writes DWG/tables.
        private void MatchTableIntoGrid()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            var peo = new PromptEntityOptions("\nSelect a crossing table (Main/Page):");
            peo.SetRejectMessage("\nEntity must be a TABLE.");
            peo.AddAllowedClass(typeof(Table), exactMatch: true);

            var sel = ed.GetEntity(peo);
            if (sel.Status != PromptStatus.OK) return;

            using (doc.LockDocument())
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var table = tr.GetObject(sel.ObjectId, OpenMode.ForRead) as Table;
                if (table == null)
                {
                    ed.WriteMessage("\n[CrossingManager] Selection was not a Table.");
                    return;
                }

                // Identify table kind (Main or Page). Lat/Long not supported here.
                var kind = _tableSync.IdentifyTable(table, tr);
                if (kind == TableSync.XingTableType.Unknown)
                {
                    ed.WriteMessage("\n[CrossingManager] Could not determine table type.");
                    return;
                }
                if (kind == TableSync.XingTableType.LatLong)
                {
                    ed.WriteMessage("\n[CrossingManager] Lat/Long tables are not supported by MATCH TABLE.");
                    return;
                }

                // Index strictly by X read from the block attribute on the cell content.
                // byX: "X1" -> (Owner, Desc, Loc, Dwg, Row)
                var byX = new Dictionary<string, (string Owner, string Desc, string Loc, string Dwg, int Row)>(StringComparer.OrdinalIgnoreCase);

                ed.WriteMessage("\n[CrossingManager] --- TABLE VALUES (parsed) ---");

                for (int r = 0; r < table.Rows.Count; r++)
                {
                    string xRaw = ReadXFromCellAttributeOnly(table, r, tr); // ← attribute on cell *content*
                    if (string.IsNullOrWhiteSpace(xRaw))
                        continue; // no X → skip row

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
                        ed.WriteMessage($"\n[CrossingManager] [U] {rec.Crossing}: grid updated from table (row {src.Row}).");
                    else
                        ed.WriteMessage($"\n[CrossingManager] [=] {rec.Crossing}: no changes needed (already matches).");
                }

                tr.Commit();
                gridCrossings.Refresh();
                _isDirty = true;

                ed.WriteMessage($"\n[CrossingManager] Match Table -> grid only (X-only): matched={matched}, updated={updated}, noMatch={noMatch}");
            }
        }
        /// Read Column A strictly from the block attribute living on the cell *content*.
        /// We try both signatures seen across releases:
        ///   GetBlockAttributeValue(string tag)
        ///   GetBlockAttributeValue(ObjectId attDefId)
        /// Returns the raw value ("X1", "X02", ...) or "" if not present.
        private static string ReadXFromCellAttributeOnly(Table table, int row, Transaction tr)
        {
            if (table == null || row < 0 || row >= table.Rows.Count) return string.Empty;

            // Tags you commonly use for the index bubble.
            string[] tags = { "CROSSING", "X", "XING", "X_NO", "XNUM", "XNUMBER", "NUMBER", "INDEX", "NO", "LABEL" };

            Cell cell = null;
            try { cell = table.Cells[row, 0]; } catch { return string.Empty; }
            if (cell == null) return string.Empty;

            foreach (var content in EnumerateCellContents(cell))
            {
                var ct = content.GetType();

                // --- (A) signature: GetBlockAttributeValue(string tag)
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
                        catch { /* try next */ }
                    }
                }

                // --- (B) signature: GetBlockAttributeValue(ObjectId attDefId)
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

            // STRICT: no fallback to block definition text/name; if we can't read the instance attribute, return "".
            return string.Empty;
        }

        /// Enumerate the cell's "Contents" collection (robust across versions).
        private static IEnumerable<object> EnumerateCellContents(Cell cell)
        {
            if (cell == null) yield break;

            var contentsProp = cell.GetType().GetProperty("Contents",
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

            var btrProp = content.GetType().GetProperty("BlockTableRecordId",
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

            try
            {
                var val = btrProp?.GetValue(content, null);
                if (val is ObjectId id) return id;
            }
            catch { }

            return ObjectId.Null;
        }

        /// Normalize "X", "X 1", "X01" -> "X1". If it doesn't look like X-digits, return compacted uppercase.
        private static string NormalizeXKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            var up = Regex.Replace(s.ToUpperInvariant(), @"\s+", "");
            // X##  -> X#
            var m = Regex.Match(up, @"^X0*(\d+)$");
            if (m.Success) return "X" + m.Groups[1].Value;
            // ##   -> X#
            m = Regex.Match(up, @"^0*(\d+)$");
            if (m.Success) return "X" + m.Groups[1].Value;
            return up;
        }

        /// Read a table cell as plain text. Treat "-" and "—" as empty.
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

        private static IEnumerable<string> EnumerateCellTextFragments(Cell cell)
        {
            if (cell == null) yield break;

            var contentsProp = cell.GetType().GetProperty("Contents",
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

            System.Collections.IEnumerable seq = null;
            try { seq = contentsProp?.GetValue(cell, null) as System.Collections.IEnumerable; } catch { }
            if (seq == null) yield break;

            foreach (var item in seq)
            {
                if (item == null) continue;
                var textProp = item.GetType().GetProperty("TextString",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                if (textProp == null) continue;

                string frag = null;
                try { frag = textProp.GetValue(item, null) as string; } catch { }
                if (!string.IsNullOrWhiteSpace(frag))
                    yield return frag;
            }
        }

        // Normalize a variety of inputs to "X#". Accepts "X 01", "X01", "01", "1", etc.
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

        private static System.Collections.Generic.IEnumerable<object> EnumerateCellContentObjects(Cell cell)
        {
            if (cell == null) yield break;

            var contentsProp = cell.GetType().GetProperty("Contents",
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

            System.Collections.IEnumerable seq = null;
            try { seq = contentsProp?.GetValue(cell, null) as System.Collections.IEnumerable; } catch { }
            if (seq == null) yield break;

            foreach (var item in seq)
                if (item != null) yield return item;
        }

        private string GenerateCrossingName(int index)
        {
            return string.Format(CultureInfo.InvariantCulture, "X{0}", index + 1);
        }
        // Treat "-" and "—" as empty placeholders for table cells

        private void ShiftCrossings(int startIndex, int delta)
        {
            if (delta == 0) return;

            for (var i = startIndex; i < _records.Count; i++)
            {
                var record = _records[i];
                var token = CrossingRecord.ParseCrossingNumber(record.Crossing);
                var prefix = ExtractPrefix(record.Crossing);
                var newNumber = Math.Max(1, token.Number + delta);
                record.Crossing = string.Format(CultureInfo.InvariantCulture, "{0}{1}", prefix, newNumber);
            }
        }

        private static string ExtractPrefix(string crossing)
        {
            if (string.IsNullOrEmpty(crossing)) return "X";

            var chars = crossing.TakeWhile(c => !char.IsDigit(c)).ToArray();
            var prefix = new string(chars);
            if (string.IsNullOrEmpty(prefix)) prefix = "X";
            return prefix;
        }

        private void RenumberSequential()
        {
            for (var i = 0; i < _records.Count; i++)
            {
                var record = _records[i];
                var prefix = ExtractPrefix(record.Crossing);
                record.Crossing = string.Format(CultureInfo.InvariantCulture, "{0}{1}", prefix, i + 1);
            }
        }

        private void GenerateXingPage()
        {
            var choices = _records
                .Select(r => r.DwgRef ?? string.Empty)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (!choices.Any())
            {
                MessageBox.Show("No DWG_REF values available.", "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var selected = PromptForChoice("Select DWG_REF", choices);
            if (string.IsNullOrEmpty(selected)) return;

            try
            {
                string actualName;
                var layoutId = _layoutUtils.CloneLayoutFromTemplate(_doc, TemplatePath, TemplateLayoutName, string.Format(CultureInfo.InvariantCulture, "X-{0}", selected), out actualName);
                _layoutUtils.SwitchToLayout(_doc, actualName);

                var locationText = BuildLocationText(selected);
                if (!string.IsNullOrEmpty(locationText))
                {
                    _layoutUtils.ReplacePlaceholderText(_doc.Database, layoutId, locationText);
                }

                var pointRes = _doc.Editor.GetPoint("\nSpecify insertion point for Crossing Page Table:");
                if (pointRes.Status != PromptStatus.OK) return;

                using (_doc.LockDocument())
                using (var tr = _doc.Database.TransactionManager.StartTransaction())
                {
                    var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                    var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);
                    _tableSync.CreateAndInsertPageTable(_doc.Database, tr, btr, pointRes.Value, selected, _records);
                    tr.Commit();
                }
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show(string.Format(CultureInfo.InvariantCulture, "Template not found: {0}", TemplatePath), "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string BuildLocationText(string dwgRef)
        {
            var record = _records.FirstOrDefault(r => string.Equals(r.DwgRef, dwgRef, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(r.Location));
            if (record == null) return string.Empty;

            string formatted;
            if (_layoutUtils.TryFormatMeridianLocation(record.Location, out formatted))
                return formatted;

            return record.Location;
        }

        private string PromptForChoice(string title, IList<string> choices)
        {
            using (var dialog = new Form())
            {
                dialog.Text = title;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.Width = 300;
                dialog.Height = 150;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;

                var combo = new ComboBox
                {
                    Dock = DockStyle.Top,
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                combo.Items.AddRange(choices.Cast<object>().ToArray());
                if (choices.Count > 0) combo.SelectedIndex = 0;

                var panel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    FlowDirection = WinFormsFlowDirection.RightToLeft,
                    Height = 40
                };

                var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 80 };
                var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 80 };
                panel.Controls.Add(ok);
                panel.Controls.Add(cancel);

                dialog.Controls.Add(combo);
                dialog.Controls.Add(panel);
                dialog.AcceptButton = ok;
                dialog.CancelButton = cancel;

                return dialog.ShowDialog(this) == DialogResult.OK ? (string)combo.SelectedItem : string.Empty;
            }
        }

        private void CreateOrUpdateLatLongTable()
        {
            var record = GetSelectedRecord();
            if (record == null)
            {
                MessageBox.Show("Select a crossing first.", "Crossing Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                        var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        var table = ent as Table;
                        if (table == null) continue;

                        if (_tableSync.IdentifyTable(table, tr) != TableSync.XingTableType.LatLong)
                            continue;

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
                _tableSync.CreateAndInsertLatLongTable(_doc.Database, tr, btr, pointRes.Value, record);
                tr.Commit();
            }
        }
    }
}
