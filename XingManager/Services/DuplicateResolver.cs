using System;
using System.Collections.Generic;
using System.Linq;
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

        public bool ResolveDuplicates(IList<CrossingRecord> records, IDictionary<ObjectId, InstanceContext> contexts)
        {
            if (records == null)
                throw new ArgumentNullException("records");

            var duplicateCandidates = BuildCandidateList(records, contexts);
            if (!duplicateCandidates.Any())
                return true;

            using (var dialog = new DuplicateResolverDialog(duplicateCandidates))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return false;

                foreach (var group in duplicateCandidates.GroupBy(c => c.CrossingKey))
                {
                    var record = records.First(r => r.CrossingKey == group.Key);
                    var selected = group.FirstOrDefault(c => c.Canonical);
                    if (selected != null)
                    {
                        // Promote selected candidate's values to the record (canonical)
                        record.CanonicalInstance = selected.ObjectId;
                        record.Crossing = selected.Crossing;
                        record.Owner = selected.Owner;
                        record.Description = selected.Description;
                        record.Location = selected.Location;
                        record.DwgRef = selected.DwgRef;
                        record.Lat = selected.Lat;
                        record.Long = selected.Long;

                        // Push chosen values into contexts for every duplicate instance
                        foreach (var candidate in group)
                        {
                            if (contexts != null)
                            {
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
                    }
                }
            }

            return true;
        }

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

                // Build UI candidates
                foreach (var objectId in record.AllInstances)
                {
                    var ctx = GetContext(contexts, objectId); // always non-null

                    // Prefer per-instance values when present; fall back to record values
                    var owner = string.IsNullOrEmpty(ctx.Owner) ? record.Owner : ctx.Owner;
                    var description = string.IsNullOrEmpty(ctx.Description) ? record.Description : ctx.Description;
                    var location = string.IsNullOrEmpty(ctx.Location) ? record.Location : ctx.Location;
                    var dwgRef = string.IsNullOrEmpty(ctx.DwgRef) ? record.DwgRef : ctx.DwgRef;
                    var lat = string.IsNullOrEmpty(ctx.Lat) ? record.Lat : ctx.Lat;
                    var lng = string.IsNullOrEmpty(ctx.Long) ? record.Long : ctx.Long;

                    list.Add(new DuplicateCandidate
                    {
                        Crossing = record.Crossing ?? string.Empty,
                        CrossingKey = record.CrossingKey,
                        ObjectId = objectId,
                        Layout = ctx.SpaceName,
                        Owner = owner ?? string.Empty,
                        Description = description ?? string.Empty,
                        Location = location ?? string.Empty,
                        DwgRef = dwgRef ?? string.Empty,
                        Lat = lat ?? string.Empty,
                        Long = lng ?? string.Empty,
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

            // Fallback context
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

        private class DuplicateResolverDialog : Form
        {
            private readonly DataGridView _grid;
            private readonly BindingSource _binding;
            private readonly List<DuplicateCandidate> _candidates;

            public DuplicateResolverDialog(List<DuplicateCandidate> candidates)
            {
                _candidates = candidates;

                Text = "Resolve Duplicate Crossings";
                Width = 1000;
                Height = 400;
                StartPosition = FormStartPosition.CenterParent;
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
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect
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
                    Width = 80
                };

                _grid.Columns.AddRange(
                    colCrossing,
                    colLayout,
                    colOwner,
                    colDescription,
                    colLocation,
                    colDwgRef,
                    colLat,
                    colLong,
                    colCanonical);

                _grid.CellContentClick += GridOnCellContentClick;

                var buttonPanel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    FlowDirection = WinFormsFlowDirection.RightToLeft,
                    Padding = new Padding(10),
                    Height = 50
                };

                var okButton = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 80 };
                var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 80 };
                buttonPanel.Controls.Add(okButton);
                buttonPanel.Controls.Add(cancelButton);

                Controls.Add(_grid);
                Controls.Add(buttonPanel);
            }

            private void GridOnCellContentClick(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0)
                    return;

                if (_grid.Columns[e.ColumnIndex] is DataGridViewCheckBoxColumn)
                {
                    var candidate = (DuplicateCandidate)_binding[e.RowIndex];

                    // Make the clicked row the canonical one within its crossing group,
                    // and propagate its values to all other candidates of the same group for preview.
                    foreach (var item in _candidates.Where(c => c.CrossingKey == candidate.CrossingKey))
                    {
                        item.Canonical = false;

                        if (!ReferenceEquals(item, candidate))
                        {
                            item.Owner = candidate.Owner;
                            item.Description = candidate.Description;
                            item.Location = candidate.Location;
                            item.DwgRef = candidate.DwgRef;
                            item.Lat = candidate.Lat;
                            item.Long = candidate.Long;
                        }
                    }

                    candidate.Canonical = true;
                    _binding.ResetBindings(false);
                }
            }
        }
    }
}
