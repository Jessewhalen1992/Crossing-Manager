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
        }

        public bool ResolveDuplicates(IList<CrossingRecord> records, IDictionary<ObjectId, InstanceContext> contexts)
        {
            if (records == null)
            {
                throw new ArgumentNullException("records");
            }

            var duplicateCandidates = BuildCandidateList(records, contexts);
            if (!duplicateCandidates.Any())
            {
                return true;
            }

            using (var dialog = new DuplicateResolverDialog(duplicateCandidates))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return false;
                }

                foreach (var group in duplicateCandidates.GroupBy(c => c.CrossingKey))
                {
                    var record = records.First(r => r.CrossingKey == group.Key);
                    var selected = group.FirstOrDefault(c => c.Canonical);
                    if (selected != null)
                    {
                        record.CanonicalInstance = selected.ObjectId;
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
                if (record.AllInstances.Count <= 1)
                {
                    continue;
                }

                ObjectId defaultCanonical = record.CanonicalInstance;
                if (defaultCanonical.IsNull)
                {
                    var firstModel = record.AllInstances
                        .Select(id => new { Id = id, Context = GetContext(contexts, id) })
                        .FirstOrDefault(t => string.Equals(t.Context.SpaceName, "Model", StringComparison.OrdinalIgnoreCase));
                    if (firstModel != null)
                    {
                        defaultCanonical = firstModel.Id;
                    }
                    else
                    {
                        defaultCanonical = record.AllInstances.First();
                    }

                    record.CanonicalInstance = defaultCanonical;
                }

                foreach (var objectId in record.AllInstances)
                {
                    var context = GetContext(contexts, objectId);
                    list.Add(new DuplicateCandidate
                    {
                        Crossing = record.Crossing ?? string.Empty,
                        CrossingKey = record.CrossingKey,
                        ObjectId = objectId,
                        Handle = objectId.Handle.ToString(),
                        Space = context.SpaceName,
                        Canonical = objectId == record.CanonicalInstance
                    });
                }
            }

            return list;
        }

        private static InstanceContext GetContext(IDictionary<ObjectId, InstanceContext> contexts, ObjectId id)
        {
            if (contexts != null && contexts.TryGetValue(id, out var ctx))
            {
                return ctx;
            }

            return new InstanceContext
            {
                ObjectId = id,
                SpaceName = "Unknown"
            };
        }

        private class DuplicateCandidate
        {
            public string Crossing { get; set; }
            public string CrossingKey { get; set; }
            public string Space { get; set; }
            public string Handle { get; set; }
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
                Width = 600;
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

                var colSpace = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.Space),
                    HeaderText = "Space",
                    ReadOnly = true,
                    Width = 120
                };

                var colHandle = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.Handle),
                    HeaderText = "Handle",
                    ReadOnly = true,
                    Width = 120
                };

                var colCanonical = new DataGridViewCheckBoxColumn
                {
                    DataPropertyName = nameof(DuplicateCandidate.Canonical),
                    HeaderText = "Canonical",
                    Width = 80
                };

                _grid.Columns.AddRange(colCrossing, colSpace, colHandle, colCanonical);
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
                {
                    return;
                }

                if (_grid.Columns[e.ColumnIndex] is DataGridViewCheckBoxColumn)
                {
                    var candidate = (DuplicateCandidate)_binding[e.RowIndex];
                    foreach (var item in _candidates.Where(c => c.CrossingKey == candidate.CrossingKey))
                    {
                        item.Canonical = false;
                    }

                    candidate.Canonical = true;
                    _binding.ResetBindings(false);
                }
            }
        }
    }
}
