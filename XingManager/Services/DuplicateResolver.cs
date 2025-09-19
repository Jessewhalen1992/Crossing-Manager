using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;                    // <-- needed for Where/Select/GroupBy
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using WinFormsFlowDirection = System.Windows.Forms.FlowDirection;
using XingManager.Models;

namespace XingManager.Services
{
    /// <summary>
    /// Presents a modal dialog that lets the user choose canonical instances for duplicates.
    /// </summary>
    public class DuplicateResolver
    {
        // ---------------------------- Context carried per BlockReference ----------------------------
        public class InstanceContext
        {
            public ObjectId ObjectId { get; set; }
            public string SpaceName { get; set; }
            public string Owner { get; set; }
            public string Description { get; set; }
            public string Location { get; set; }
            public string DwgRef { get; set; }
            public string Lat { get; set; }
            public string Long { get; set; }
        }

        /// <summary>
        /// Show the dialog if duplicates exist; on OK, set one canonical per crossing group and
        /// write the chosen values back into the in-memory records/contexts. The caller handles DB apply.
        /// </summary>
        public bool ResolveDuplicates(IList<CrossingRecord> records, IDictionary<ObjectId, InstanceContext> contexts)
        {
            if (records == null)
                throw new ArgumentNullException("records");

            var duplicateCandidates = BuildCandidateList(records, contexts);
            if (!duplicateCandidates.Any())
                return true; // nothing to resolve

            var duplicateGroups = duplicateCandidates
                .GroupBy(c => c.CrossingKey, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.ToList())
                .Where(group => group.Count > 1)
                .ToList();

            for (var i = 0; i < duplicateGroups.Count; i++)
            {
                var group = duplicateGroups[i];
                var displayName = !string.IsNullOrWhiteSpace(group[0].Crossing)
                    ? group[0].Crossing
                    : group[0].CrossingKey;

                using (var dialog = new DuplicateResolverDialog(group, displayName, i + 1, duplicateGroups.Count))
                {
                    if (dialog.ShowDialog() != DialogResult.OK)
                        return false;
                }

                var selected = group.FirstOrDefault(c => c.Canonical);
                if (selected == null)
                    continue;

                var record = records.First(r => string.Equals(r.CrossingKey, group[0].CrossingKey, StringComparison.OrdinalIgnoreCase));

                // Promote selected candidate's values to the record (canonical snapshot)
                record.CanonicalInstance = selected.ObjectId;
                record.Crossing = selected.Crossing;
                record.Owner = selected.Owner;
                record.Description = selected.Description;
                record.Location = selected.Location;
                record.DwgRef = selected.DwgRef;
                record.Lat = selected.Lat;
                record.Long = selected.Long;

                // Write the chosen values into every instance context of the group
                foreach (var candidate in group)
                {
                    if (contexts == null)
                        continue;

                    InstanceContext ctx;
                    if (contexts.TryGetValue(candidate.ObjectId, out ctx) && ctx != null)
                    {
                        ctx.Owner = selected.Owner;
                        ctx.Description = selected.Description;
                        ctx.Location = selected.Location;
                        ctx.DwgRef = selected.DwgRef;
                        ctx.Lat = selected.Lat;
                        ctx.Long = selected.Long;
                    }
                }
            }

            return true;
        }

        // ---------------------------- Build candidate rows for the dialog ----------------------------
        private static List<DuplicateCandidate> BuildCandidateList(IEnumerable<CrossingRecord> records, IDictionary<ObjectId, InstanceContext> contexts)
        {
            var list = new List<DuplicateCandidate>();

            foreach (var record in records)
            {
                if (record.AllInstances == null || record.AllInstances.Count <= 1)
                    continue;

                // Choose a default canonical if one isn't set (prefer Model space)
                ObjectId defaultCanonical = record.CanonicalInstance;
                if (defaultCanonical.IsNull)
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

                // Build UI candidates (use per-instance values only so differences are visible)
                foreach (var objectId in record.AllInstances)
                {
                    var ctx = GetContext(contexts, objectId);

                    list.Add(new DuplicateCandidate
                    {
                        Crossing = record.Crossing ?? string.Empty,
                        CrossingKey = record.CrossingKey,
                        ObjectId = objectId,
                        Layout = ctx.SpaceName ?? "Unknown",
                        Owner = ctx.Owner ?? string.Empty,
                        Description = ctx.Description ?? string.Empty,
                        Location = ctx.Location ?? string.Empty,
                        DwgRef = ctx.DwgRef ?? string.Empty,
                        Lat = ctx.Lat ?? string.Empty,
                        Long = ctx.Long ?? string.Empty,
                        Canonical = objectId == record.CanonicalInstance
                    });
                }
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
                SpaceName = "Unknown",
                Owner = string.Empty,
                Description = string.Empty,
                Location = string.Empty,
                DwgRef = string.Empty,
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
            public string Lat { get; set; }
            public string Long { get; set; }
            public ObjectId ObjectId { get; set; }
            public bool Canonical { get; set; }
        }

        // ---------------------------- Inner dialog (WINFORMS) ----------------------------
        private class DuplicateResolverDialog : Form
        {
            private static Point? _lastDialogLocation;
            private readonly DataGridView _grid;
            private readonly BindingSource _binding;
            private readonly List<DuplicateCandidate> _candidates;
            private readonly string _crossingLabel;

            public DuplicateResolverDialog(List<DuplicateCandidate> candidates, string crossingLabel, int groupIndex, int groupCount)
            {
                if (candidates == null)
                    throw new ArgumentNullException("candidates");

                _candidates = candidates;
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
                _binding.DataSource = _candidates;

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
                    colLocation, colDwgRef, colLat, colLong, colCanonical);

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

                    var candidate = (DuplicateCandidate)_binding[e.RowIndex];

                    // Toggle only within this CrossingKey group
                    foreach (var item in _candidates.Where(c => string.Equals(c.CrossingKey, candidate.CrossingKey, StringComparison.OrdinalIgnoreCase)))
                    {
                        item.Canonical = ReferenceEquals(item, candidate);
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

            private static string BuildDialogTitle(int groupIndex, int groupCount)
            {
                if (groupCount <= 0)
                    return "Resolve Duplicate Crossings";

                groupIndex = Math.Max(1, groupIndex);

                return string.Format("Resolve Duplicate Crossings ({0} of {1})", groupIndex, groupCount);
            }
        }
    }
}
