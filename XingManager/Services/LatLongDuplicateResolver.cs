using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using XingManager.Models;
using WinFormsFlowDirection = System.Windows.Forms.FlowDirection;

namespace XingManager.Services
{
    /// <summary>
    /// Handles duplicate resolution for LAT/LONG values coming from blocks and LAT/LONG tables.
    /// Adds a deterministic "find & replace" write-back so existing tables are updated immediately.
    /// </summary>
    public class LatLongDuplicateResolver
    {
        public bool ResolveDuplicates(
            IList<CrossingRecord> records,
            IDictionary<ObjectId, DuplicateResolver.InstanceContext> contexts)
        {
            if (records == null)
                throw new ArgumentNullException(nameof(records));

            var candidates = BuildCandidateList(records, contexts);
            if (!candidates.Any())
                return true;

            var groups = candidates
                .GroupBy(c => c.CrossingKey, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.ToList())
                .Where(group => group.Count > 1)
                .ToList();

            // Track which crossings we actually changed so we can push to DWG tables
            var changed = new List<(string CrossingKey, LatLongCandidate Canonical)>();

            for (var i = 0; i < groups.Count; i++)
            {
                var group = groups[i];
                var displayName = !string.IsNullOrWhiteSpace(group[0].Crossing)
                    ? group[0].Crossing
                    : group[0].CrossingKey;

                using (var dialog = new LatLongDuplicateResolverDialog(group, displayName, i + 1, groups.Count))
                {
                    if (ModelessDialogRunner.ShowDialog(dialog) != DialogResult.OK)
                        return false;
                }

                var selected = group.FirstOrDefault(c => c.Canonical);
                if (selected == null)
                    continue;

                // Promote to the CrossingRecord (authoritative in-memory snapshot)
                var record = records.First(r => string.Equals(r.CrossingKey, group[0].CrossingKey, StringComparison.OrdinalIgnoreCase));

                record.Lat = selected.Lat;
                record.Long = selected.Long;
                record.Zone = selected.Zone;
                if (!string.IsNullOrWhiteSpace(selected.DwgRef))
                    record.DwgRef = selected.DwgRef;

                // Keep existing LatLongSources in sync
                if (record.LatLongSources != null && record.LatLongSources.Count > 0)
                {
                    foreach (var source in record.LatLongSources)
                    {
                        source.Lat = selected.Lat;
                        source.Long = selected.Long;
                        source.Zone = selected.Zone;
                        if (!string.IsNullOrWhiteSpace(selected.DwgRef))
                            source.DwgRef = selected.DwgRef;
                    }
                }

                // Keep live contexts (blocks/table instances) in sync
                foreach (var instanceId in record.AllInstances ?? Enumerable.Empty<ObjectId>())
                {
                    if (contexts != null && contexts.TryGetValue(instanceId, out var ctx) && ctx != null)
                    {
                        ctx.Lat = selected.Lat;
                        ctx.Long = selected.Long;
                        ctx.Zone = selected.Zone;
                    }
                }

                changed.Add((record.CrossingKey, selected));
            }

            // ------------------------------------------------------------------
            // Critical part: immediately push chosen LAT/LONG into DWG tables.
            // This prevents the UI from "snapping back" when it refreshes from DWG.
            // ------------------------------------------------------------------
            if (changed.Count > 0)
            {
                try
                {
                    ApplyLatLongChoicesToDrawingTables(records, changed);
                }
                catch
                {
                    // Do not fail duplicate resolution if DWG write-back had an issue.
                    // The normal TableSync pass will still try to push changes later.
                }
            }

            return true;
        }

        // =============================================================================
        // DWG "find & replace" for LAT/LONG tables
        // =============================================================================

        /// <summary>
        /// Writes selected LAT/LONG/ZONE into the actual drawing tables:
        ///  (1) exact source rows we know about (TableId/RowIndex),
        ///  (2) any other LAT/LONG tables whose "ID" cell matches the crossing key.
        /// </summary>
        private static void ApplyLatLongChoicesToDrawingTables(
            IList<CrossingRecord> allRecords,
            IEnumerable<(string CrossingKey, LatLongCandidate Canonical)> changes)
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager?.MdiActiveDocument;
            if (doc == null)
                return;

            var byKey = (allRecords ?? new List<CrossingRecord>())
                .Where(r => r != null)
                .ToDictionary(r => r.CrossingKey, r => r, StringComparer.OrdinalIgnoreCase);

            // Group by table for the (1) exact source rows case
            var rowsByTable = new Dictionary<ObjectId, List<(int Row, CrossingRecord Rec, LatLongCandidate Sel)>>();

            foreach (var (key, sel) in changes)
            {
                if (!byKey.TryGetValue(key, out var rec) || rec == null)
                    continue;

                foreach (var src in rec.LatLongSources ?? Enumerable.Empty<CrossingRecord.LatLongSource>())
                {
                    if (src == null || src.TableId.IsNull || !src.TableId.IsValid || src.RowIndex < 0)
                        continue;

                    if (!rowsByTable.TryGetValue(src.TableId, out var list))
                    {
                        list = new List<(int, CrossingRecord, LatLongCandidate)>();
                        rowsByTable[src.TableId] = list;
                    }
                    list.Add((src.RowIndex, rec, sel));
                }
            }

            using (doc.LockDocument())
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                // (1) Write directly into known source rows
                foreach (var kvp in rowsByTable)
                {
                    var table = tr.GetObject(kvp.Key, OpenMode.ForWrite, false, true) as Table;
                    if (table == null) continue;

                    var hasExtended = table.Columns.Count >= 6;
                    var zoneCol = hasExtended ? 2 : -1;
                    var latCol = hasExtended ? 3 : 2;
                    var lngCol = hasExtended ? 4 : 3;
                    var dwgCol = hasExtended ? 5 : -1;

                    foreach (var (row, rec, sel) in kvp.Value)
                    {
                        if (row < 0 || row >= table.Rows.Count) continue;

                        SafeSetCell(table, row, latCol, sel.Lat);
                        SafeSetCell(table, row, lngCol, sel.Long);
                        if (zoneCol >= 0) SafeSetCell(table, row, zoneCol, rec.ZoneLabel ?? sel.Zone ?? string.Empty);
                        if (dwgCol >= 0) SafeSetCell(table, row, dwgCol, rec.DwgRef ?? sel.DwgRef ?? string.Empty);
                    }

                    TryRefresh(table);
                }

                // (2) Sweep all tables: replace wherever ID == crossing key
                //     (covers legacy tables / rows not recorded in LatLongSources)
                var db = doc.Database;
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    foreach (ObjectId entId in btr)
                    {
                        var table = tr.GetObject(entId, OpenMode.ForWrite, false, true) as Table;
                        if (table == null) continue;

                        // Only 4- or 6-col layouts are treated as LAT/LONG candidates
                        var cols = table.Columns.Count;
                        if (cols != 4 && cols != 6) continue;

                        // Use the same start-row logic as TableSync
                        var dataStart = TableSync.FindLatLongDataStartRow(table);
                        if (dataStart <= 0) dataStart = 1; // fallback for header-less

                        var hasExtended = cols >= 6;
                        var zoneCol = hasExtended ? 2 : -1;
                        var latCol = hasExtended ? 3 : 2;
                        var lngCol = hasExtended ? 4 : 3;
                        var dwgCol = hasExtended ? 5 : -1;

                        bool anyRowChanged = false;

                        for (int row = dataStart; row < table.Rows.Count; row++)
                        {
                            // Robust key read (handles text, blocks, attribute tags, mtext, etc.)
                            var rawKey = TableSync.ResolveCrossingKey(table, row, 0);
                            var normKey = TableSync.NormalizeKeyForLookup(rawKey); // "X4", "X10", ...
                            if (string.IsNullOrEmpty(normKey)) continue;

                            // Do we have a change for this key?
                            var change = changes.FirstOrDefault(c =>
                                string.Equals(c.CrossingKey, normKey, StringComparison.OrdinalIgnoreCase));
                            if (string.IsNullOrEmpty(change.CrossingKey)) continue;

                            // Resolve record to pick Zone/DWG if present
                            byKey.TryGetValue(change.CrossingKey, out var rec);

                            SafeSetCell(table, row, latCol, change.Canonical.Lat);
                            SafeSetCell(table, row, lngCol, change.Canonical.Long);
                            if (zoneCol >= 0) SafeSetCell(table, row, zoneCol, rec?.ZoneLabel ?? change.Canonical.Zone ?? string.Empty);
                            if (dwgCol >= 0) SafeSetCell(table, row, dwgCol, rec?.DwgRef ?? change.Canonical.DwgRef ?? string.Empty);

                            anyRowChanged = true;
                        }

                        if (anyRowChanged)
                            TryRefresh(table);
                    }
                }

                tr.Commit();
            }
        }

        private static void SafeSetCell(Table table, int row, int col, string value)
        {
            if (table == null || row < 0 || col < 0) return;
            if (row >= table.Rows.Count || col >= table.Columns.Count) return;

            try { table.Cells[row, col].TextString = value ?? string.Empty; }
            catch { /* swallow; we never want dialog to fail */ }
        }

        private static void TryRefresh(Table table)
        {
            if (table == null) return;

            // Prefer RecomputeTableBlock(true) when available; else GenerateLayout()
            var mi = table.GetType().GetMethod("RecomputeTableBlock", new[] { typeof(bool) });
            if (mi != null)
            {
                try { mi.Invoke(table, new object[] { true }); return; } catch { }
            }
            try { table.GenerateLayout(); } catch { }
        }

        // =============================================================================
        // Original candidate building & dialog
        // =============================================================================

        private static List<LatLongCandidate> BuildCandidateList(
            IEnumerable<CrossingRecord> records,
            IDictionary<ObjectId, DuplicateResolver.InstanceContext> contexts)
        {
            var list = new List<LatLongCandidate>();

            foreach (var record in records)
            {
                if (record == null)
                    continue;

                var recordCandidates = new List<LatLongCandidate>();
                var normalizedRecordLat = Normalize(record.Lat);
                var normalizedRecordLong = Normalize(record.Long);
                var normalizedRecordZone = Normalize(record.Zone);

                var instances = record.AllInstances ?? new List<ObjectId>();
                foreach (var objectId in instances)
                {
                    var ctx = GetContext(contexts, objectId);
                    if (ctx.IgnoreForDuplicates)
                        continue;

                    var candidate = new LatLongCandidate
                    {
                        Crossing = record.Crossing ?? record.CrossingKey,
                        CrossingKey = record.CrossingKey,
                        SourceType = "Block",
                        Source = BuildBlockSourceLabel(ctx),
                        Description = ctx.Description ?? record.Description ?? string.Empty,
                        Lat = ctx.Lat ?? string.Empty,
                        Long = ctx.Long ?? string.Empty,
                        Zone = ctx.Zone ?? string.Empty,
                        DwgRef = ctx.DwgRef ?? record.DwgRef ?? string.Empty,
                        ObjectId = objectId
                    };

                    if (MatchesRecord(candidate, normalizedRecordLat, normalizedRecordLong, normalizedRecordZone))
                        candidate.Canonical = true;

                    recordCandidates.Add(candidate);
                }

                var latSources = record.LatLongSources ?? new List<CrossingRecord.LatLongSource>();
                foreach (var source in latSources)
                {
                    var sourceLabel = !string.IsNullOrWhiteSpace(source.SourceLabel)
                        ? source.SourceLabel
                        : "LAT/LONG Table";

                    var candidate = new LatLongCandidate
                    {
                        Crossing = record.Crossing ?? record.CrossingKey,
                        CrossingKey = record.CrossingKey,
                        SourceType = "Table",
                        Source = sourceLabel,
                        Description = source.Description ?? record.Description ?? string.Empty,
                        Lat = source.Lat ?? string.Empty,
                        Long = source.Long ?? string.Empty,
                        Zone = source.Zone ?? string.Empty,
                        DwgRef = source.DwgRef ?? record.DwgRef ?? string.Empty,
                        TableId = source.TableId,
                        RowIndex = source.RowIndex
                    };

                    if (MatchesRecord(candidate, normalizedRecordLat, normalizedRecordLong, normalizedRecordZone))
                        candidate.Canonical = true;

                    recordCandidates.Add(candidate);
                }

                recordCandidates = recordCandidates
                    .Where(c => !string.IsNullOrWhiteSpace(Normalize(c.Lat)) ||
                                !string.IsNullOrWhiteSpace(Normalize(c.Long)) ||
                                !string.IsNullOrWhiteSpace(Normalize(c.Zone)))
                    .ToList();

                if (recordCandidates.Count <= 1 || !RequiresResolution(recordCandidates))
                    continue;

                if (!recordCandidates.Any(c => c.Canonical))
                    recordCandidates[0].Canonical = true;

                list.AddRange(recordCandidates);
            }

            return list;
        }

        private static bool MatchesRecord(LatLongCandidate candidate, string recordLat, string recordLong, string recordZone)
        {
            if (string.IsNullOrWhiteSpace(recordLat) &&
                string.IsNullOrWhiteSpace(recordLong) &&
                string.IsNullOrWhiteSpace(recordZone))
                return false;

            return string.Equals(Normalize(candidate.Lat), recordLat, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(Normalize(candidate.Long), recordLong, StringComparison.OrdinalIgnoreCase) &&
                   (string.IsNullOrWhiteSpace(recordZone) ||
                    string.Equals(Normalize(candidate.Zone), recordZone, StringComparison.OrdinalIgnoreCase));
        }

        private static DuplicateResolver.InstanceContext GetContext(
            IDictionary<ObjectId, DuplicateResolver.InstanceContext> contexts,
            ObjectId id)
        {
            if (contexts != null && contexts.TryGetValue(id, out var ctx) && ctx != null)
                return ctx;

            return new DuplicateResolver.InstanceContext
            {
                ObjectId = id,
                Crossing = string.Empty,
                SpaceName = "Unknown",
                Owner = string.Empty,
                Description = string.Empty,
                Location = string.Empty,
                DwgRef = string.Empty,
                Zone = string.Empty,
                Lat = string.Empty,
                Long = string.Empty
            };
        }

        private static bool RequiresResolution(List<LatLongCandidate> candidates)
        {
            if (candidates == null || candidates.Count <= 1)
                return false;

            var signatures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var candidate in candidates)
            {
                signatures.Add(BuildCandidateSignature(candidate));
                if (signatures.Count > 1)
                    return true;
            }

            return false;
        }

        private static string BuildCandidateSignature(LatLongCandidate candidate)
        {
            if (candidate == null)
                return string.Empty;

            return string.Join("|",
                Normalize(candidate.Lat),
                Normalize(candidate.Long),
                Normalize(candidate.Zone));
        }

        private static string Normalize(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? string.Empty
                : value.Trim();
        }

        private static string BuildBlockSourceLabel(DuplicateResolver.InstanceContext ctx)
        {
            if (ctx == null)
                return string.Empty;

            var layout = string.IsNullOrWhiteSpace(ctx.SpaceName) ? "Unknown" : ctx.SpaceName.Trim();
            return string.Format(CultureInfo.InvariantCulture, "Block ({0})", layout);
        }

        private class LatLongCandidate
        {
            public string Crossing { get; set; }
            public string CrossingKey { get; set; }
            public string SourceType { get; set; }
            public string Source { get; set; }
            public string Description { get; set; }
            public string Lat { get; set; }
            public string Long { get; set; }
            public string Zone { get; set; }
            public string DwgRef { get; set; }
            public ObjectId ObjectId { get; set; }
            public ObjectId TableId { get; set; }
            public int RowIndex { get; set; }
            public bool Canonical { get; set; }
        }

        // --------------------------- Dialog ---------------------------

        private class LatLongDuplicateResolverDialog : Form
        {
            private static Point? _lastDialogLocation;
            private readonly DataGridView _grid;
            private readonly BindingSource _binding;
            private readonly List<LatLongCandidate> _candidates;
            private readonly List<DisplayCandidate> _displayCandidates;
            private readonly string _crossingLabel;

            public LatLongDuplicateResolverDialog(List<LatLongCandidate> candidates, string crossingLabel, int groupIndex, int groupCount)
            {
                if (candidates == null)
                    throw new ArgumentNullException(nameof(candidates));

                _candidates = candidates;
                _displayCandidates = BuildDisplayCandidates(_candidates);
                _crossingLabel = string.IsNullOrWhiteSpace(crossingLabel) ? "Crossing" : crossingLabel;

                Text = BuildDialogTitle(groupIndex, groupCount);
                Width = 800;
                Height = 360;
                StartPosition = _lastDialogLocation.HasValue
                    ? FormStartPosition.Manual
                    : FormStartPosition.CenterParent;
                if (_lastDialogLocation.HasValue)
                    Location = _lastDialogLocation.Value;
                MinimizeBox = false;
                MaximizeBox = false;
                FormBorderStyle = FormBorderStyle.FixedDialog;

                _binding = new BindingSource { DataSource = _displayCandidates };

                _grid = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    AutoGenerateColumns = false,
                    DataSource = _binding,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    ReadOnly = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    EditMode = DataGridViewEditMode.EditOnEnter
                };

                var colType = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DisplayCandidate.SourceType),
                    HeaderText = "Type",
                    ReadOnly = true,
                    Width = 120
                };

                var colDescription = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DisplayCandidate.Description),
                    HeaderText = "Description",
                    ReadOnly = true,
                    Width = 260
                };

                var colLat = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DisplayCandidate.Lat),
                    HeaderText = "Lat",
                    ReadOnly = true,
                    Width = 120
                };

                var colLong = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DisplayCandidate.Long),
                    HeaderText = "Long",
                    ReadOnly = true,
                    Width = 120
                };

                var colCanonical = new DataGridViewCheckBoxColumn
                {
                    DataPropertyName = nameof(DisplayCandidate.Canonical),
                    HeaderText = "Canonical",
                    Width = 80
                };

                _grid.Columns.AddRange(colType, colDescription, colLat, colLong, colCanonical);
                _grid.CellContentClick += GridOnCellContentClick;

                var headerLabel = new Label
                {
                    Dock = DockStyle.Top,
                    Padding = new Padding(10),
                    Height = 40,
                    Text = string.Format(
                        CultureInfo.InvariantCulture,
                        "Select the canonical LAT/LONG for {0} ({1} of {2}).",
                        _crossingLabel,
                        groupIndex,
                        groupCount)
                };

                var buttonPanel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    FlowDirection = WinFormsFlowDirection.RightToLeft,
                    Padding = new Padding(10),
                    Height = 50
                };

                var okButton = new Button { Text = "OK", Width = 80 };
                okButton.Click += OkButtonOnClick;

                var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 80 };
                cancelButton.Click += CancelButtonOnClick;
                buttonPanel.Controls.Add(okButton);
                buttonPanel.Controls.Add(cancelButton);

                Controls.Add(_grid);
                Controls.Add(buttonPanel);
                Controls.Add(headerLabel);

                AcceptButton = okButton;
                CancelButton = cancelButton;
            }

            protected override void OnFormClosed(FormClosedEventArgs e)
            {
                _lastDialogLocation = Location;
                base.OnFormClosed(e);
            }

            private void GridOnCellContentClick(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0)
                    return;

                if (_grid.Columns[e.ColumnIndex] is DataGridViewCheckBoxColumn)
                {
                    _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    _grid.EndEdit();

                    var candidate = (DisplayCandidate)_binding[e.RowIndex];

                    foreach (var item in _displayCandidates.Where(c => string.Equals(c.CrossingKey, candidate.CrossingKey, StringComparison.OrdinalIgnoreCase)))
                    {
                        item.SetCanonical(ReferenceEquals(item, candidate));
                    }

                    _grid.Refresh();
                }
            }

            private void OkButtonOnClick(object sender, EventArgs e)
            {
                var groups = _displayCandidates
                    .GroupBy(c => c.CrossingKey, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var group in groups)
                {
                    if (group.Count(c => c.Canonical) != 1)
                    {
                        MessageBox.Show(
                            string.Format(CultureInfo.InvariantCulture, "Please select exactly one canonical LAT/LONG for {0}.", _crossingLabel),
                            "Resolve Duplicate LAT/LONG",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        DialogResult = DialogResult.None;
                        return;
                    }
                }

                DialogResult = DialogResult.OK;
                Close();
            }

            private void CancelButtonOnClick(object sender, EventArgs e)
            {
                DialogResult = DialogResult.Cancel;
                Close();
            }

            private static string BuildDialogTitle(int groupIndex, int groupCount)
            {
                if (groupCount <= 0)
                    return "Resolve Duplicate LAT/LONG";

                groupIndex = Math.Max(1, groupIndex);
                return string.Format(CultureInfo.InvariantCulture, "Resolve Duplicate LAT/LONG ({0} of {1})", groupIndex, groupCount);
            }

            private static List<DisplayCandidate> BuildDisplayCandidates(IEnumerable<LatLongCandidate> candidates)
            {
                if (candidates == null)
                    return new List<DisplayCandidate>();

                return candidates
                    .GroupBy(c => BuildCandidateSignature(c), StringComparer.OrdinalIgnoreCase)
                    .Select(g => new DisplayCandidate(g.ToList()))
                    .ToList();
            }

            private class DisplayCandidate
            {
                private readonly List<LatLongCandidate> _members;

                public DisplayCandidate(List<LatLongCandidate> members)
                {
                    if (members == null || members.Count == 0)
                        throw new ArgumentException(nameof(members));

                    _members = members;
                }

                private LatLongCandidate Representative => _members[0];

                public string CrossingKey => Representative.CrossingKey;
                public string SourceType => Representative.SourceType;
                public string Source => Representative.Source;
                public string Description => Representative.Description;
                public string Lat => Representative.Lat;
                public string Long => Representative.Long;

                public bool Canonical
                {
                    get => _members.Any(m => m.Canonical);
                    set => SetCanonical(value);
                }

                public void SetCanonical(bool isCanonical)
                {
                    foreach (var member in _members)
                        member.Canonical = isCanonical;
                }
            }
        }
    }
}
