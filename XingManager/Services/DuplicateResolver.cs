using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;                    // <-- needed for Where/Select/GroupBy
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using WinFormsFlowDirection = System.Windows.Forms.FlowDirection;
using XingManager.Models;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace XingManager.Services
{
    /// <summary>
    /// Presents a modal dialog that lets the user choose canonical instances for duplicates.
    /// </summary>
    public class DuplicateResolver
    {
        // ---------------------------- Context carried per BlockReference ----------------------------
        // Inside DuplicateResolver.cs
        public class InstanceContext
        {
            public ObjectId ObjectId { get; set; }
            public string Crossing { get; set; }
            public string SpaceName { get; set; }
            public string Owner { get; set; }
            public string Description { get; set; }
            public string Location { get; set; }
            public string DwgRef { get; set; }
            public string Zone { get; set; }
            public string Lat { get; set; }
            public string Long { get; set; }

            // New flag: true if this instance is inside a table or on a paper-space layout.
            // Such instances should be updated silently but not shown in the duplicate resolver UI.
            public bool IgnoreForDuplicates { get; set; }
        }


        /// <summary>
        /// Show the dialog if duplicates exist; on OK, set one canonical per crossing group and
        /// write the chosen values back into the in‑memory records/contexts. The caller handles DB apply.
        /// </summary>
        public bool ResolveDuplicates(IList<CrossingRecord> records, IDictionary<ObjectId, InstanceContext> contexts)
        {
            if (records == null)
                throw new ArgumentNullException("records");

            var ed = AutoCADApp.DocumentManager?.MdiActiveDocument?.Editor;
            // Build the flat list shown in the dialog (skips IgnoreForDuplicates instances)
            var duplicateCandidates = BuildCandidateList(records, contexts);
            Logger.Info(ed, $"dup_blocks candidates={duplicateCandidates.Count}");
            if (!duplicateCandidates.Any())
                return true; // nothing to resolve

            // Group by CrossingKey, only show groups that actually differ
            var duplicateGroups = duplicateCandidates
                .GroupBy(c => c.CrossingKey, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.ToList())
                .Where(group => group.Count > 1)
                .ToList();

            Logger.Info(ed, $"dup_blocks groups={duplicateGroups.Count}");
            var resolvedGroups = 0;

            for (var i = 0; i < duplicateGroups.Count; i++)
            {
                var group = duplicateGroups[i];
                var displayName = !string.IsNullOrWhiteSpace(group[0].Crossing)
                    ? group[0].Crossing
                    : group[0].CrossingKey;

                using (var dialog = new DuplicateResolverDialog(group, displayName, i + 1, duplicateGroups.Count))
                {
                    if (ModelessDialogRunner.ShowDialog(dialog) != DialogResult.OK)
                        return false; // user canceled
                }

                // The one the user ticked
                var selected = group.FirstOrDefault(c => c.Canonical);
                if (selected == null)
                {
                    Logger.Debug(ed, $"dup_blocks group crossing={displayName} canonical=none members={group.Count}");
                    continue;
                }

                resolvedGroups++;
                var canonicalHandle = !selected.ObjectId.IsNull ? selected.ObjectId.Handle.ToString() : "null";
                Logger.Debug(ed, $"dup_blocks group crossing={displayName} canonical={canonicalHandle} members={group.Count}");

                // Find & update the record (canonical snapshot)
                var record = records.First(r =>
                    string.Equals(r.CrossingKey, group[0].CrossingKey, StringComparison.OrdinalIgnoreCase));

                if (!selected.ObjectId.IsNull)
                    record.CanonicalInstance = selected.ObjectId;

                record.Crossing = selected.Crossing;
                record.Owner = selected.Owner;
                record.Description = selected.Description;
                record.Location = selected.Location;
                record.DwgRef = selected.DwgRef;
                record.Zone = selected.Zone;
                record.Lat = selected.Lat;     // <— IMPORTANT
                record.Long = selected.Long;    // <— IMPORTANT

                // Keep every instance context synced (even those not shown in the dialog)
                foreach (var instanceId in record.AllInstances)
                {
                    if (contexts != null && contexts.TryGetValue(instanceId, out var ctx) && ctx != null)
                    {
                        ctx.Crossing = selected.Crossing;
                        ctx.Owner = selected.Owner;
                        ctx.Description = selected.Description;
                        ctx.Location = selected.Location;
                        ctx.DwgRef = selected.DwgRef;
                        ctx.Zone = selected.Zone;
                        ctx.Lat = selected.Lat;   // <— IMPORTANT
                        ctx.Long = selected.Long;  // <— IMPORTANT
                    }
                }
            }

            Logger.Info(ed, $"dup_blocks summary groups={duplicateGroups.Count} resolved={resolvedGroups} skipped={duplicateGroups.Count - resolvedGroups}");

            return true;
        }

        // ---------------------------- Build candidate rows for the dialog ----------------------------
        // Build a flat list of duplicate candidates to display in the UI
        private static List<DuplicateCandidate> BuildCandidateList(IEnumerable<CrossingRecord> records, IDictionary<ObjectId, InstanceContext> contexts)
        {
            var list = new List<DuplicateCandidate>();

            foreach (var record in records)
            {
                // Only consider records with more than one instance
                var instanceCount = record.AllInstances?.Count ?? 0;
                if (instanceCount <= 1)
                    continue;

                // Choose a default canonical if one isn't set (prefer Model space)
                ObjectId defaultCanonical = record.CanonicalInstance;
                if (instanceCount > 0 && defaultCanonical.IsNull)
                {
                    var firstModel = record.AllInstances
                        .Select(id => new { Id = id, Context = GetContext(contexts, id) })
                        .FirstOrDefault(t => string.Equals(t.Context.SpaceName, "Model", StringComparison.OrdinalIgnoreCase));

                    if (firstModel != null)
                        defaultCanonical = firstModel.Id;
                    else
                        defaultCanonical = record.AllInstances[0];

                    record.CanonicalInstance = defaultCanonical;
                }

                // Build per-instance UI candidates; skip those flagged as IgnoreForDuplicates
                var recordCandidates = new List<DuplicateCandidate>();
                if (instanceCount > 0)
                {
                    foreach (var objectId in record.AllInstances)
                    {
                        var ctx = GetContext(contexts, objectId);
                        if (ctx.IgnoreForDuplicates)
                            continue; // table cell or paper-space block: update later but don't show

                        var crossing = !string.IsNullOrWhiteSpace(ctx.Crossing)
                            ? ctx.Crossing
                            : record.Crossing ?? string.Empty;

                        recordCandidates.Add(new DuplicateCandidate
                        {
                            Crossing = crossing,
                            CrossingKey = record.CrossingKey,
                            ObjectId = objectId,
                            Layout = ctx.SpaceName ?? "Unknown",
                            Owner = ctx.Owner ?? string.Empty,
                            Description = ctx.Description ?? string.Empty,
                            Location = ctx.Location ?? string.Empty,
                            DwgRef = ctx.DwgRef ?? string.Empty,
                            Zone = ctx.Zone ?? string.Empty,
                            Lat = ctx.Lat ?? string.Empty,
                            Long = ctx.Long ?? string.Empty,
                            Canonical = objectId == record.CanonicalInstance
                        });
                    }
                }

                // If all instances are ignored or there’s no visible discrepancy, skip this record
                if (!recordCandidates.Any() || !RequiresResolution(recordCandidates))
                    continue;

                if (!recordCandidates.Any(c => c.Canonical))
                    recordCandidates[0].Canonical = true;

                list.AddRange(recordCandidates);
            }

            return list;
        }

        private static InstanceContext GetContext(IDictionary<ObjectId, InstanceContext> contexts, ObjectId id)
        {
            if (contexts != null)
            {
                InstanceContext ctx;
                if (contexts.TryGetValue(id, out ctx) && ctx != null)
                    return ctx;
            }

            // Fallback context (empty values)
            return new InstanceContext
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

        // ---------------------------- Inner model used by the dialog grid ----------------------------
        private class DuplicateCandidate
        {
            public string Crossing { get; set; }
            public string CrossingKey { get; set; }
            public string Layout { get; set; }
            public string Owner { get; set; }
            public string Description { get; set; }
            public string Location { get; set; }
            public string DwgRef { get; set; }
            public string Zone { get; set; }
            public string Lat { get; set; }
            public string Long { get; set; }
            public ObjectId ObjectId { get; set; }
            public bool Canonical { get; set; }
            public string ZoneLabel
            {
                get
                {
                    if (string.IsNullOrWhiteSpace(Zone))
                        return string.Empty;

                    return string.Format(CultureInfo.InvariantCulture, "ZONE {0}", Zone.Trim());
                }
            }
        }

        private static bool RequiresResolution(List<DuplicateCandidate> candidates)
        {
            if (candidates == null || candidates.Count <= 1)
                return false;

            var signatures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var candidate in candidates)
            {
                var signature = BuildCandidateSignature(candidate);
                signatures.Add(signature);

                if (signatures.Count > 1)
                    return true;
            }

            return false;
        }

        private static string BuildCandidateSignature(DuplicateCandidate candidate)
        {
            if (candidate == null)
                return string.Empty;

            return string.Join(
                "|",
                NormalizeAttribute(candidate.Owner),
                NormalizeAttribute(candidate.Description),
                NormalizeAttribute(candidate.Location),
                NormalizeAttribute(candidate.DwgRef),
                NormalizeAttribute(candidate.Zone),
                NormalizeAttribute(candidate.Lat),
                NormalizeAttribute(candidate.Long));
        }

        private static string NormalizeAttribute(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? string.Empty
                : value.Trim();
        }

        // ---------------------------- Inner dialog (WINFORMS) ----------------------------
        private class DuplicateResolverDialog : Form
        {
            private static Point? _lastDialogLocation;
            private readonly DataGridView _grid;
            private readonly BindingSource _binding;
            private readonly List<DuplicateCandidate> _candidates;
            private readonly List<DisplayCandidate> _displayCandidates;
            private readonly string _crossingLabel;

            public DuplicateResolverDialog(List<DuplicateCandidate> candidates, string crossingLabel, int groupIndex, int groupCount)
            {
                if (candidates == null)
                    throw new ArgumentNullException("candidates");

                _candidates = candidates;
                _displayCandidates = BuildDisplayCandidates(_candidates);
                _crossingLabel = string.IsNullOrWhiteSpace(crossingLabel) ? "Crossing" : crossingLabel;

                Text = BuildDialogTitle(groupIndex, groupCount);
                Width = 1000;
                Height = 400;
                StartPosition = _lastDialogLocation.HasValue
                    ? FormStartPosition.Manual
                    : FormStartPosition.CenterParent;
                if (_lastDialogLocation.HasValue)
                    Location = _lastDialogLocation.Value;
                MinimizeBox = false;
                MaximizeBox = false;
                FormBorderStyle = FormBorderStyle.FixedDialog;

                _binding = new BindingSource();
                _binding.DataSource = _displayCandidates;

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

                var colCrossing = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.Crossing),
                    HeaderText = "Crossing",
                    ReadOnly = true,
                    Width = 80
                };
                var colLayout = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.Layout),
                    HeaderText = "Layout",
                    ReadOnly = true,
                    Width = 120
                };
                var colOwner = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.Owner),
                    HeaderText = "Owner",
                    ReadOnly = true,
                    Width = 120
                };
                var colDescription = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.Description),
                    HeaderText = "Description",
                    ReadOnly = true,
                    Width = 200
                };
                var colLocation = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.Location),
                    HeaderText = "Location",
                    ReadOnly = true,
                    Width = 200
                };
                var colZone = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.ZoneLabel),
                    HeaderText = "Zone",
                    ReadOnly = true,
                    Width = 80
                };
                var colDwgRef = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.DwgRef),
                    HeaderText = "DWG_REF",
                    ReadOnly = true,
                    Width = 120
                };
                var colLat = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.Lat),
                    HeaderText = "Lat",
                    ReadOnly = true,
                    Width = 80
                };
                var colLong = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.Long),
                    HeaderText = "Long",
                    ReadOnly = true,
                    Width = 80
                };
                var colCanonical = new DataGridViewCheckBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.Canonical),
                    HeaderText = "Canonical",
                    Width = 80,
                    ThreeState = false
                };

                _grid.Columns.AddRange(
                    colCrossing, colLayout, colOwner, colDescription,
                    colLocation, colZone, colDwgRef, colLat, colLong, colCanonical);

                _grid.CellContentClick += GridOnCellContentClick;

                var headerLabel = new Label
                {
                    Dock = DockStyle.Top,
                    Padding = new Padding(10),
                    Height = 40,
                    Text = string.Format("Select the canonical instance for {0} ({1} of {2}).", _crossingLabel, groupIndex, groupCount)
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

                    // Toggle only within this CrossingKey group
                    foreach (var item in _displayCandidates.Where(c => string.Equals(c.CrossingKey, candidate.CrossingKey, StringComparison.OrdinalIgnoreCase)))
                    {
                        item.SetCanonical(ReferenceEquals(item, candidate));
                    }

                    _binding.ResetBindings(false);
                }
            }

            private void OkButtonOnClick(object sender, EventArgs e)
            {
                var chosen = _candidates.Count(c => c.Canonical);
                if (chosen != 1)
                {
                    MessageBox.Show(
                        "Please select exactly one canonical for " + _crossingLabel + ".",
                        "Resolve Duplicate Crossings",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    this.DialogResult = DialogResult.None; // prevent close
                    return;
                }

                this.DialogResult = DialogResult.OK;
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
                    return "Resolve Duplicate Crossings";

                groupIndex = Math.Max(1, groupIndex);

                return string.Format("Resolve Duplicate Crossings ({0} of {1})", groupIndex, groupCount);
            }

            private static List<DisplayCandidate> BuildDisplayCandidates(IEnumerable<DuplicateCandidate> candidates)
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
                private readonly List<DuplicateCandidate> _members;

                public DisplayCandidate(List<DuplicateCandidate> members)
                {
                    if (members == null || members.Count == 0)
                        throw new ArgumentException("members");

                    _members = members;
                }

                private DuplicateCandidate Representative
                {
                    get { return _members[0]; }
                }

                public string Crossing
                {
                    get { return Representative.Crossing; }
                }

                public string CrossingKey
                {
                    get { return Representative.CrossingKey; }
                }

                public string Layout
                {
                    get { return Representative.Layout; }
                }

                public string Owner
                {
                    get { return Representative.Owner; }
                }

                public string Description
                {
                    get { return Representative.Description; }
                }

                public string Location
                {
                    get { return Representative.Location; }
                }

                public string Zone
                {
                    get { return Representative.Zone; }
                }

                public string DwgRef
                {
                    get { return Representative.DwgRef; }
                }

                public string ZoneLabel
                {
                    get { return Representative.ZoneLabel; }
                }

                public string Lat
                {
                    get { return Representative.Lat; }
                }

                public string Long
                {
                    get { return Representative.Long; }
                }

                public bool Canonical
                {
                    get { return _members.Any(m => m.Canonical); }
                    set { SetCanonical(value); }
                }

                public void SetCanonical(bool isCanonical)
                {
                    foreach (var member in _members)
                        member.Canonical = false;

                    if (isCanonical)
                        Representative.Canonical = true;
                }
            }
        }
    }
}
