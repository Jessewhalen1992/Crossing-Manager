using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using XingManager.Models;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

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

            // Flag used by the duplicate-resolver UIs to hide noisy instances (e.g., table-cell blocks or paper-space copies).
            // Note: some workflows (like comparing Crossing Tables vs. blocks) may intentionally override this to false
            // so the table-row values can participate in resolution.
            public bool IgnoreForDuplicates { get; set; }

            // True if this block is physically located inside any AutoCAD Table cell.
            // Kept separate from IgnoreForDuplicates so other resolvers (e.g., LAT/LONG) can always exclude
            // table-cell blocks even if IgnoreForDuplicates was overridden elsewhere.
            public bool IsTableInstance { get; set; }
        }


        /// <summary>
        /// Show the dialog if duplicates exist; on OK, set one canonical per crossing group and
        /// write the chosen values back into the in-memory records/contexts. The caller handles DB apply.
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

            try
            {
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
                        Logger.Info(ed, $"dup_blocks group crossing={displayName} canonical=none members={group.Count}");
                        continue;
                    }

                    resolvedGroups++;
                    var canonicalHandle = !selected.ObjectId.IsNull ? selected.ObjectId.Handle.ToString() : "null";
                    Logger.Info(ed, $"dup_blocks group crossing={displayName} canonical={canonicalHandle} members={group.Count}");

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
                    // NOTE: Zone/Lat/Long are resolved by LatLongDuplicateResolver (separate workflow).
                    // Do not overwrite them here when resolving crossing-table/block duplicates.

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
                            // Zone/Lat/Long are handled by LatLongDuplicateResolver.
                        }
                    }

                    // Immediately push the chosen canonical values back to the drawing blocks so the user
                    // sees the result right away (even if they only ran Scan).
                    TryApplyToDrawing(record);
                }
            }
            finally
            {
                // If the resolver made any changes, also sync table text so Scan-only workflows
                // don't immediately re-trigger the same discrepancies.
                if (resolvedGroups > 0)
                {
                    TryUpdateTables(records);
                }
            }

            Logger.Info(ed, $"dup_blocks summary groups={duplicateGroups.Count} resolved={resolvedGroups} skipped={duplicateGroups.Count - resolvedGroups}");

            return true;
        }

        /// <summary>
        /// Best-effort table synchronization after a duplicate resolution.
        /// We intentionally keep this here (instead of in the Form) so that Scan-only
        /// flows still propagate the chosen canonical values to table text.
        /// </summary>
        private static void TryUpdateTables(IList<CrossingRecord> records)
        {
            if (records == null || records.Count == 0)
                return;

            try
            {
                var doc = AutoCADApp.DocumentManager?.MdiActiveDocument;
                if (doc == null)
                    return;

                var tableSync = new TableSync(new TableFactory());
                // NOTE: LAT/LONG tables are handled by LatLongDuplicateResolver. Don't touch LAT/LONG tables here.
                tableSync.UpdateCrossingTables(doc, records);
            }
            catch (System.Exception ex)
            {
                var ed = AutoCADApp.DocumentManager?.MdiActiveDocument?.Editor;
                Logger.Warn(ed, $"dup_blocks update_tables failed: {ex.Message}");
            }
        }

        // ---------------------------- Build candidate rows for the dialog ----------------------------
        // Build a flat list of duplicate candidates to display in the UI
        private static List<DuplicateCandidate> BuildCandidateList(IEnumerable<CrossingRecord> records, IDictionary<ObjectId, InstanceContext> contexts)
        {
            var list = new List<DuplicateCandidate>();
            var ed = AutoCADApp.DocumentManager?.MdiActiveDocument?.Editor;

            Logger.Info(ed, "=== BuildCandidateList START ===");
            
            if (records == null)
            {
                Logger.Info(ed, "BuildCandidateList: records is NULL, returning empty list");
                return list;
            }

            var recordCount = records.Count();
            Logger.Info(ed, $"BuildCandidateList: Processing {recordCount} records");
            
            int recordsWithInstances = 0;
            int recordsWithoutInstances = 0;

            foreach (var record in records)
            {
                if (record == null)
                {
                    Logger.Info(ed, "  Found NULL record, skipping");
                    continue;
                }

                var instances = record.AllInstances ?? new List<ObjectId>();
                var instanceCount = instances.Count;
                
                if (instanceCount > 0)
                    recordsWithInstances++;
                else
                    recordsWithoutInstances++;
                
                Logger.Info(ed, $"BuildCandidates: {record.Crossing ?? record.CrossingKey} - instances={instanceCount}, tableSources={record.CrossingTableSources?.Count ?? 0}");
                
                if (instanceCount <= 0)
                {
                    Logger.Info(ed, $"  SKIPPED - no instances");
                    continue;
                }

                // Choose a default canonical if one isn't set (prefer Model space)
                ObjectId defaultCanonical = record.CanonicalInstance;
                if (defaultCanonical.IsNull || !defaultCanonical.IsValid)
                {
                    var firstModel = instances
                        .Select(id => new { Id = id, Context = GetContext(contexts, id) })
                        .FirstOrDefault(t => string.Equals(t.Context.SpaceName, "Model", StringComparison.OrdinalIgnoreCase));

                    if (firstModel != null)
                        defaultCanonical = firstModel.Id;
                    else
                        defaultCanonical = instances[0];

                    record.CanonicalInstance = defaultCanonical;
                }

                // Build candidates from:
                //   - the in-memory record snapshot (UI)
                //   - every non-ignored block instance (including table-row overrides when present)
                var recordCandidates = new List<DuplicateCandidate>();

                // UI / record snapshot candidate (lets us detect mismatches between table text and a single block instance)
                var uiCandidate = new DuplicateCandidate
                {
                    Crossing = record.Crossing ?? record.CrossingKey,
                    CrossingKey = record.CrossingKey,
                    ObjectId = ObjectId.Null,
                    Layout = "UI",
                    Owner = record.Owner ?? string.Empty,
                    Description = record.Description ?? string.Empty,
                    Location = record.Location ?? string.Empty,
                    DwgRef = record.DwgRef ?? string.Empty,
                    Zone = record.Zone ?? string.Empty,
                    Lat = record.Lat ?? string.Empty,
                    Long = record.Long ?? string.Empty,
                    Canonical = false
                };
                recordCandidates.Add(uiCandidate);

                // Per-instance candidates
                // CRITICAL FIX: ApplyCrossingTableOverrides sets IgnoreForDuplicates=false for table bubble blocks
                // that were successfully matched to table row data. Those blocks should participate in duplicate
                // resolution (with their table-derived values from the context).
                // We ONLY skip instances where IgnoreForDuplicates is still true (unmatched table blocks or paper space).
                int blockCandidatesAdded = 0;
                foreach (var objectId in instances)
                {
                    var ctx = GetContext(contexts, objectId);
                    
                    Logger.Info(ed, $"  Instance {objectId.Handle}: IsTableInstance={ctx.IsTableInstance}, IgnoreForDuplicates={ctx.IgnoreForDuplicates}, Space={ctx.SpaceName}");
                    
                    // Skip only if explicitly marked to ignore
                    if (ctx.IgnoreForDuplicates)
                    {
                        Logger.Info(ed, $"    SKIPPED (IgnoreForDuplicates=true)");
                        continue;
                    }

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
                    blockCandidatesAdded++;
                }
                
                Logger.Info(ed, $"  Added {blockCandidatesAdded} block candidates");


                // Add candidates sourced from crossing tables (table row text)
                // These represent the values from table cells (columns B, C, D, etc.) for each X# found in column A.
                int tableCandidatesAdded = 0;
                if (record.CrossingTableSources != null && record.CrossingTableSources.Count > 0)
                {
                    foreach (var src in record.CrossingTableSources)
                    {
                        if (src == null)
                            continue;

                        recordCandidates.Add(new DuplicateCandidate
                        {
                            Crossing = record.Crossing ?? record.CrossingKey,
                            CrossingKey = record.CrossingKey,
                            ObjectId = ObjectId.Null,
                            Layout = string.IsNullOrWhiteSpace(src.SourceLabel) ? "TABLE" : src.SourceLabel,
                            Owner = src.HasOwner ? (src.Owner ?? string.Empty) : (record.Owner ?? string.Empty),
                            Description = src.Description ?? string.Empty,
                            Location = src.HasLocation ? (src.Location ?? string.Empty) : (record.Location ?? string.Empty),
                            DwgRef = src.HasDwgRef ? (src.DwgRef ?? string.Empty) : (record.DwgRef ?? string.Empty),
                            Zone = record.Zone ?? string.Empty,
                            Lat = record.Lat ?? string.Empty,
                            Long = record.Long ?? string.Empty,
                            Canonical = false
                        });
                        tableCandidatesAdded++;
                    }
                }
                
                Logger.Info(ed, $"  Added {tableCandidatesAdded} table candidates, total={recordCandidates.Count}");

                // If we only have the UI candidate (no visible instances), there's nothing to compare
                if (recordCandidates.Count <= 1)
                {
                    Logger.Info(ed, $"  SKIPPED - only {recordCandidates.Count} candidate (need 2+)");
                    continue;
                }

                // If there's no visible discrepancy, skip this record
                if (!RequiresResolution(recordCandidates))
                {
                    Logger.Info(ed, $"  SKIPPED - no discrepancy (all values identical)");
                    continue;
                }
                
                Logger.Info(ed, $"  INCLUDED - has discrepancies, adding {recordCandidates.Count} candidates to list");

                // Ensure exactly one canonical default.
                // Prefer the existing CanonicalInstance if present; otherwise prefer Model space; otherwise UI.
                if (!recordCandidates.Any(c => c.Canonical))
                {
                    var firstModel = recordCandidates
                        .FirstOrDefault(c => !c.ObjectId.IsNull && string.Equals(c.Layout, "Model", StringComparison.OrdinalIgnoreCase));

                    if (firstModel != null)
                        firstModel.Canonical = true;
                    else
                    {
                        var firstInstance = recordCandidates.FirstOrDefault(c => !c.ObjectId.IsNull);
                        if (firstInstance != null)
                            firstInstance.Canonical = true;
                        else
                            uiCandidate.Canonical = true;
                    }
                }

                // If the default canonical comes from a table-instance (values read from table text) and it differs
                // from the UI snapshot, default to UI so we don't silently accept a manual table edit.
                var canonical = recordCandidates.FirstOrDefault(c => c.Canonical);
                if (canonical != null && uiCandidate != null)
                {
                    if (!canonical.ObjectId.IsNull && contexts != null &&
                        contexts.TryGetValue(canonical.ObjectId, out var canonCtx) &&
                        canonCtx != null && canonCtx.IsTableInstance)
                    {
                        var canonSig = BuildCandidateSignature(canonical);
                        var uiSig = BuildCandidateSignature(uiCandidate);

                        if (!string.Equals(canonSig, uiSig, StringComparison.OrdinalIgnoreCase))
                        {
                            canonical.Canonical = false;
                            uiCandidate.Canonical = true;
                        }
                    }
                }

                list.AddRange(recordCandidates);
            }

            Logger.Info(ed, $"BuildCandidateList SUMMARY: {recordsWithInstances} records WITH instances, {recordsWithoutInstances} records WITHOUT instances");
            Logger.Info(ed, $"=== BuildCandidateList END: returning {list.Count} total candidates ===");
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

            // This resolver is for Crossing-table / block attribute fields only:
            //   CROSSING / OWNER / DESCRIPTION / LOCATION / DWG_REF
            // Zone/Lat/Long are resolved separately by LatLongDuplicateResolver.
            return string.Join(
                "|",
                NormalizeAttribute(candidate.Crossing),
                NormalizeAttribute(candidate.Owner),
                NormalizeAttribute(candidate.Description),
                NormalizeAttribute(candidate.Location),
                NormalizeAttribute(candidate.DwgRef));
        }

        private static string NormalizeAttribute(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            var v = value.Trim();

            // Treat a lone dash as "blank" (tables often use "-" placeholders)
            if (v == "-" || v == "–" || v == "—")
                return string.Empty;

            return v;
        }

        // ---------------------------- Write-back helpers ----------------------------

        private const string LatLongDictionaryKey = "XING2_LATLNG";
        private const string LatLongKeyLat = "LAT";
        private const string LatLongKeyLong = "LONG";
        private const string LatLongKeyZone = "ZONE";

        /// <summary>
        /// Immediately writes the resolved canonical values back to every block instance for the record.
        /// This keeps the drawing consistent even when the user only ran "Scan" (no Apply).
        /// </summary>
        private static void TryApplyToDrawing(CrossingRecord record)
        {
            if (record == null || record.AllInstances == null || record.AllInstances.Count == 0)
                return;

            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null)
                return;

            try
            {
                using (doc.LockDocument())
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    foreach (var id in record.AllInstances)
                    {
                        if (id == ObjectId.Null || !id.IsValid)
                            continue;

                        try
                        {
                            var br = tr.GetObject(id, OpenMode.ForWrite, false) as BlockReference;
                            if (br == null || br.IsErased)
                                continue;

                            SetAttributeValue(tr, br, "CROSSING", record.Crossing);
                            SetAttributeValue(tr, br, "OWNER", record.Owner);
                            SetAttributeValue(tr, br, "DESCRIPTION", record.Description);
                            SetAttributeValue(tr, br, "LOCATION", record.Location);
                            SetAttributeValue(tr, br, "DWG_REF", record.DwgRef);

                            // Zone/Lat/Long are handled separately by LatLongDuplicateResolver; don't overwrite here.
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception aex)
                        {
                            // Common non-fatal cases when working with stale ObjectIds.
                            if (aex.ErrorStatus == ErrorStatus.WasErased ||
                                aex.ErrorStatus == ErrorStatus.InvalidInput ||
                                aex.ErrorStatus == ErrorStatus.NullObjectId)
                                continue;
                        }
                        catch
                        {
                            // ignore per-instance failures
                        }
                    }

                    tr.Commit();
                }
            }
            catch
            {
                // ignore - writeback is best-effort
            }
        }

        private static void SetAttributeValue(Transaction tr, BlockReference br, string tag, string value)
        {
            if (tr == null || br == null || string.IsNullOrWhiteSpace(tag))
                return;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                if (attId == ObjectId.Null)
                    continue;

                var attRef = tr.GetObject(attId, OpenMode.ForWrite, false) as AttributeReference;
                if (attRef == null)
                    continue;

                if (string.Equals(attRef.Tag, tag, StringComparison.OrdinalIgnoreCase))
                {
                    attRef.TextString = value ?? string.Empty;
                    return;
                }
            }
        }

        private static void SetLatLong(Transaction tr, BlockReference br, string lat, string lng, string zone)
        {
            if (tr == null || br == null)
                return;

            try
            {
                if (br.ExtensionDictionary == ObjectId.Null)
                    br.CreateExtensionDictionary();

                if (br.ExtensionDictionary == ObjectId.Null)
                    return;

                var extDict = tr.GetObject(br.ExtensionDictionary, OpenMode.ForWrite, false) as DBDictionary;
                if (extDict == null)
                    return;

                Xrecord xr;
                if (extDict.Contains(LatLongDictionaryKey))
                {
                    xr = tr.GetObject(extDict.GetAt(LatLongDictionaryKey), OpenMode.ForWrite, false) as Xrecord;
                }
                else
                {
                    xr = new Xrecord();
                    extDict.SetAt(LatLongDictionaryKey, xr);
                    tr.AddNewlyCreatedDBObject(xr, true);
                }

                if (xr == null)
                    return;

                xr.Data = new ResultBuffer(
                    new TypedValue((int)DxfCode.Text, LatLongKeyLat),
                    new TypedValue((int)DxfCode.Text, lat ?? string.Empty),
                    new TypedValue((int)DxfCode.Text, LatLongKeyLong),
                    new TypedValue((int)DxfCode.Text, lng ?? string.Empty),
                    new TypedValue((int)DxfCode.Text, LatLongKeyZone),
                    new TypedValue((int)DxfCode.Text, zone ?? string.Empty)
                );
            }
            catch
            {
                // ignore
            }
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
                    FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
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
