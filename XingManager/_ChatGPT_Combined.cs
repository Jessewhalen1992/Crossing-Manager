// AUTO-GENERATED MERGE FILE FOR REVIEW
// Generated: 2026-01-21 20:44:37

/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Commands.cs
/////////////////////////////////////////////////////////////////////

using Autodesk.AutoCAD.Runtime;

namespace XingManager
{
    public class Commands
    {
        [CommandMethod("XINGFORM")]
        public void ShowForm()
        {
            var app = XingManagerApp.Instance;
            if (app == null)
            {
                return;
            }

            app.ShowPalette();
        }

        [CommandMethod("XINGAPPLY")]
        public void ApplyChanges()
        {
            var app = XingManagerApp.Instance;
            if (app == null)
            {
                return;
            }

            var form = app.GetOrCreateForm();
            form?.ApplyToDrawing();
        }

        [CommandMethod("XINGPAGE")]
        public void GeneratePage()
        {
            var app = XingManagerApp.Instance;
            if (app == null)
            {
                return;
            }

            var form = app.GetOrCreateForm();
            form?.GenerateXingPageFromCommand();
        }

        [CommandMethod("XINGLATROW")]
        public void CreateLatLongRow()
        {
            var app = XingManagerApp.Instance;
            if (app == null)
            {
                return;
            }

            var form = app.GetOrCreateForm();
            form?.CreateLatLongRowFromCommand();
        }

        [CommandMethod("XINGREN")]
        public void Renumber()
        {
            var app = XingManagerApp.Instance;
            if (app == null)
            {
                return;
            }

            var form = app.GetOrCreateForm();
            form?.RenumberSequentiallyFromCommand();
        }

        [CommandMethod("XINGRNCPL")]
        public void CreateRncPolyline()
        {
            var app = XingManagerApp.Instance;
            if (app == null)
            {
                return;
            }

            var form = app.GetOrCreateForm();
            form?.AddRncPolylineFromCommand();
        }
    }
}

/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\DebugCommands.cs
/////////////////////////////////////////////////////////////////////

using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using XingManager.Services;

namespace XingManager
{
    public class DebugCommands
    {
        // Add this to your main commands class file
        [CommandMethod("XING_DUMP_CELL")]
        public void XingDumpCell()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            // Prompt for the table
            var peo = new PromptEntityOptions("\nSelect a TABLE:");
            peo.SetRejectMessage("\nMust be a Table object.");
            peo.AddAllowedClass(typeof(Table), true);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            // Prompt for row index
            var pioRow = new PromptIntegerOptions("\nRow index (0-based):");
            pioRow.AllowNegative = false;
            var pirRow = ed.GetInteger(pioRow);
            if (pirRow.Status != PromptStatus.OK) return;
            int row = pirRow.Value;

            // Prompt for column index
            var pioCol = new PromptIntegerOptions("\nColumn index (0-based):");
            pioCol.AllowNegative = false;
            var pirCol = ed.GetInteger(pioCol);
            if (pirCol.Status != PromptStatus.OK) return;
            int col = pirCol.Value;

            ed.WriteMessage($"\n[Dump] Cell ({row},{col})\n");

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var table = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table;
                if (table == null) return;

                if (row >= table.Rows.Count || col >= table.Columns.Count)
                {
                    ed.WriteMessage("\nError: Row or column index is out of bounds.");
                    return;
                }

                var cell = table.Cells[row, col];

                ed.WriteMessage("\n--- Cell Properties ---\n");
                ed.WriteMessage($"Cell.DataType: {cell.DataType}\n");
                ed.WriteMessage($"Cell.Value Type: {(cell.Value?.GetType().FullName ?? "null")}\n");
                ed.WriteMessage($"Cell.Value: {(cell.Value?.ToString() ?? "null")}\n");
                ed.WriteMessage($"Cell.TextString: {cell.TextString}\n");

                if (cell.Value is ObjectId blockId && !blockId.IsNull)
                {
                    ed.WriteMessage("\n--- Block Attributes ---\n");
                    var blockRef = tr.GetObject(blockId, OpenMode.ForRead) as BlockReference;
                    if (blockRef != null)
                    {
                        ed.WriteMessage($"Block Name: {blockRef.Name}\n");
                        if (blockRef.AttributeCollection.Count == 0)
                        {
                            ed.WriteMessage("-> No attributes found on this block.\n");
                        }
                        foreach (ObjectId attId in blockRef.AttributeCollection)
                        {
                            var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                            if (attRef != null)
                            {
                                ed.WriteMessage($"-> TAG: '{attRef.Tag}' = VALUE: '{attRef.TextString}'\n");
                            }
                        }
                    }
                    else
                    {
                        ed.WriteMessage("-> Value is an ObjectId, but it's not a BlockReference.\n");
                    }
                }
                else
                {
                    ed.WriteMessage("\n-> Cell.Value is NOT a block ObjectId.\n");
                }
                tr.Commit();
            }
        }

        [CommandMethod("XING_FIXBORDERS")]
        public void FixAllTableBorders()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
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
                        var t = tr.GetObject(entId, OpenMode.ForRead) as Table;
                        if (t == null) continue;
                        t.UpgradeOpen();

                        int rows = t.Rows.Count;
                        int cols = t.Columns.Count;

                        int dataStart = 0;
                        try { dataStart = TableSync.FindLatLongDataStartRow(t); if (dataStart < 0) dataStart = 0; } catch { dataStart = 0; }

                        bool IsHeading(int r)
                        {
                            try
                            {
                                var s = (t.Cells[r, 0]?.TextString ?? string.Empty).Trim().ToUpperInvariant();
                                if (s.Contains("CROSSING INFORMATION")) return true;
                            }
                            catch { }
                            return (dataStart > 0 && r == dataStart - 1);
                        }

                        for (int r = 0; r < rows; r++)
                            for (int c = 0; c < cols; c++)
                            {
                                bool heading = IsHeading(r);
                                TrySetGridVisibility(t, r, c, GridLineType.HorizontalTop, !heading);
                                TrySetGridVisibility(t, r, c, GridLineType.VerticalRight, !heading);
                                TrySetGridVisibility(t, r, c, GridLineType.HorizontalBottom, true);
                                TrySetGridVisibility(t, r, c, GridLineType.VerticalLeft, !heading);
                            }

                        try { t.GenerateLayout(); } catch { }
                        try { t.RecordGraphicsModified(true); } catch { }
                    }
                }
                tr.Commit();
            }

            try { doc.Editor?.Regen(); } catch { }
        }

        private static MethodInfo _debugSetGridVisibilityBool;
        private static MethodInfo _debugSetGridVisibilityEnum;
        private static Type _debugVisibilityType;

        private static void TrySetGridVisibility(Table t, int row, int col, GridLineType line, bool visible)
        {
            if (t == null) return;

            try
            {
                var tableType = t.GetType();
                if (_debugSetGridVisibilityBool == null)
                {
                    _debugSetGridVisibilityBool = tableType.GetMethod(
                        "SetGridVisibility",
                        new[] { typeof(int), typeof(int), typeof(GridLineType), typeof(bool) });
                }

                if (_debugSetGridVisibilityBool != null)
                {
                    _debugSetGridVisibilityBool.Invoke(t, new object[] { row, col, line, visible });
                    return;
                }

                if (_debugVisibilityType == null)
                {
                    _debugVisibilityType = tableType.Assembly.GetType("Autodesk.AutoCAD.DatabaseServices.Visibility");
                }
                if (_debugVisibilityType == null)
                    return;

                if (_debugSetGridVisibilityEnum == null)
                {
                    _debugSetGridVisibilityEnum = tableType.GetMethod(
                        "SetGridVisibility",
                        new[] { typeof(int), typeof(int), typeof(GridLineType), _debugVisibilityType });
                }

                if (_debugSetGridVisibilityEnum == null)
                    return;

                object enumValue = null;
                try
                {
                    var field = _debugVisibilityType.GetField(visible ? "Visible" : "Invisible", BindingFlags.Public | BindingFlags.Static);
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
                        enumValue = Enum.Parse(_debugVisibilityType, visible ? "Visible" : "Invisible");
                    }
                    catch
                    {
                        return;
                    }
                }

                _debugSetGridVisibilityEnum.Invoke(t, new[] { (object)row, (object)col, (object)line, enumValue });
            }
            catch
            {
                // best effort compatibility for different AutoCAD versions
            }
        }
    }
}

/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Interop\MapOdShim.cs
/////////////////////////////////////////////////////////////////////

#if ENABLE_MAP_OD
using Autodesk.Gis.Map;
#endif

namespace XingManager.Interop
{
    /// <summary>
    /// Optional shim for Map Object Data related helpers. In the base
    /// implementation this is a no-op to keep the core plug-in free from
    /// Map 3D dependencies.
    /// </summary>
    public static class MapOdShim
    {
#if ENABLE_MAP_OD
        // Placeholder for future Map Object Data integration.
#endif
    }
}

/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Models\CrossingRecord.cs
/////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;

namespace XingManager.Models
{
    /// <summary>
    /// Represents one canonical crossing record tracked by the manager.
    /// </summary>
    public class CrossingRecord
    {
        public string Crossing { get; set; }

        public string Owner { get; set; }

        public string Description { get; set; }

        public string Location { get; set; }

        public string DwgRef { get; set; }

        public string Lat { get; set; }

        public string Long { get; set; }

        public string Zone { get; set; }

        public string ZoneLabel
        {
            get
            {
                var zone = Zone?.Trim();
                if (string.IsNullOrEmpty(zone))
                {
                    return string.Empty;
                }

                return string.Format(CultureInfo.InvariantCulture, "ZONE {0}", zone);
            }
        }

        public List<ObjectId> AllInstances { get; } = new List<ObjectId>();

        public List<CrossingTableSource> CrossingTableSources { get; } = new List<CrossingTableSource>();


        public List<LatLongSource> LatLongSources { get; } = new List<LatLongSource>();

        public ObjectId CanonicalInstance { get; set; }

        public string CrossingKey
        {
            get { return (Crossing ?? string.Empty).Trim().ToUpperInvariant(); }
        }

        public void SetFrom(CrossingRecord other)
        {
            if (other == null)
            {
                return;
            }

            Crossing = other.Crossing;
            Owner = other.Owner;
            Description = other.Description;
            Location = other.Location;
            DwgRef = other.DwgRef;
            Lat = other.Lat;
            Long = other.Long;
            Zone = other.Zone;
        }

        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture, "{0} - {1}", Crossing, Description);
        }

        public static int CompareByCrossing(CrossingRecord left, CrossingRecord right)
        {
            if (left == null && right == null)
            {
                return 0;
            }

            if (left == null)
            {
                return -1;
            }

            if (right == null)
            {
                return 1;
            }

            return CompareCrossingKeys(left.Crossing, right.Crossing);
        }

        public static int CompareCrossingKeys(string left, string right)
        {
            var leftKey = ParseCrossingNumber(left);
            var rightKey = ParseCrossingNumber(right);

            var numberCompare = leftKey.Number.CompareTo(rightKey.Number);
            if (numberCompare != 0)
            {
                return numberCompare;
            }

            return string.Compare(leftKey.Suffix, rightKey.Suffix, StringComparison.OrdinalIgnoreCase);
        }

        public static CrossingToken ParseCrossingNumber(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return new CrossingToken(0, string.Empty);
            }

            var trimmed = value.Trim();
            var digits = new string(trimmed.Where(char.IsDigit).ToArray());
            int number;
            if (!int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out number))
            {
                number = 0;
            }

            var suffix = new string(trimmed.Where(c => !char.IsDigit(c)).ToArray());
            return new CrossingToken(number, suffix);
        }

        public struct CrossingToken
        {
            public CrossingToken(int number, string suffix)
            {
                Number = number;
                Suffix = suffix ?? string.Empty;
            }

            public int Number { get; private set; }

            public string Suffix { get; private set; }
        }

        

public class CrossingTableSource
{
    public string SourceLabel { get; set; }

    public string Owner { get; set; }

    public string Description { get; set; }

    public string Location { get; set; }

    public string DwgRef { get; set; }

    public bool HasOwner { get; set; }

    public bool HasLocation { get; set; }

    public bool HasDwgRef { get; set; }

    public ObjectId TableId { get; set; }

    public int RowIndex { get; set; }
}

public class LatLongSource
        {
            public string SourceLabel { get; set; }

            public string Description { get; set; }

            public string Lat { get; set; }

            public string Long { get; set; }

            public string Zone { get; set; }

            public string DwgRef { get; set; }

            public ObjectId TableId { get; set; }

            public int RowIndex { get; set; }
        }
    }
}


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Properties\AssemblyInfo.cs
/////////////////////////////////////////////////////////////////////

using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("XingManager")]
[assembly: AssemblyDescription("AutoCAD Map 3D Crossing Manager")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("XingManager")]
[assembly: AssemblyCopyright("Copyright Â© 2024")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

[assembly: ComVisible(false)]

[assembly: Guid("b7edb6ef-ff6f-4cb5-9390-1f8367b3c2b2")]

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\CommandLogger.cs
/////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Text;
using Autodesk.AutoCAD.EditorInput;

namespace XingManager.Services
{
    internal static class CommandLogger
    {
        private static readonly object _sync = new object();
        private static readonly string _logFilePath = InitializeLogFilePath();

        private static string InitializeLogFilePath()
        {
            string rootDirectory = null;

            try
            {
                rootDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (string.IsNullOrWhiteSpace(rootDirectory))
                {
                    rootDirectory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                }
            }
            catch
            {
                rootDirectory = null;
            }

            if (string.IsNullOrWhiteSpace(rootDirectory))
            {
                try
                {
                    rootDirectory = Path.GetTempPath();
                }
                catch
                {
                    rootDirectory = ".";
                }
            }

            string folder;
            try
            {
                folder = Path.Combine(rootDirectory, "CrossingManager", "Logs");
            }
            catch
            {
                folder = rootDirectory;
            }

            try
            {
                Directory.CreateDirectory(folder);
            }
            catch
            {
                folder = rootDirectory;
            }

            var fileName = $"CrossingManager_{DateTime.Now:yyyyMMdd}.log";
            try
            {
                return Path.Combine(folder, fileName);
            }
            catch
            {
                return Path.Combine(rootDirectory, "CrossingManager.log");
            }
        }

        private static void AppendLine(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
                return;

            var timestamped = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [CrossingManager] {message}";

            lock (_sync)
            {
                try
                {
                    File.AppendAllText(_logFilePath, timestamped + Environment.NewLine, Encoding.UTF8);
                }
                catch
                {
                    // Swallow logging failures to avoid impacting command execution.
                }
            }
        }

        public static void Log(string message)
        {
            AppendLine(message);
        }

        public static void Log(Editor editor, string message, bool alsoToCommandBar = false)
        {
            AppendLine(message);

            if (!alsoToCommandBar || editor == null)
                return;

            Logger.Info(editor, message);
        }

        public static string LogFilePath => _logFilePath;
    }
}


/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\DuplicateResolver.cs
/////////////////////////////////////////////////////////////////////

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
            if (v == "-" || v == "â€“" || v == "â€”")
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

/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\LatLongDuplicateResolver.cs
/////////////////////////////////////////////////////////////////////

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
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;

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

            var ed = AutoCADApp.DocumentManager?.MdiActiveDocument?.Editor;
            var candidates = BuildCandidateList(records, contexts);
            Logger.Info(ed, $"dup_latlong candidates={candidates.Count}");
            if (!candidates.Any())
                return true;

            var groups = candidates
                .GroupBy(c => c.CrossingKey, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.ToList())
                .Where(group => group.Count > 1)
                .ToList();

            Logger.Info(ed, $"dup_latlong groups={groups.Count}");
            var autoResolvedCount = 0;

            // Track which crossings we actually changed so we can push to DWG tables
            var changed = new List<(string CrossingKey, LatLongCandidate Canonical)>();

            for (var i = 0; i < groups.Count; i++)
            {
                var group = groups[i];
                var displayName = !string.IsNullOrWhiteSpace(group[0].Crossing)
                    ? group[0].Crossing
                    : group[0].CrossingKey;

                // NEW: auto-resolve if all duplicate candidates have identical LAT and LONG
                bool allSame = group.All(c =>
                    string.Equals(Normalize(c.Lat), Normalize(group[0].Lat), StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(Normalize(c.Long), Normalize(group[0].Long), StringComparison.OrdinalIgnoreCase));

                if (allSame)
                {
                    // Mark the first candidate as canonical, clear the rest
                    group[0].Canonical = true;
                    for (int j = 1; j < group.Count; j++)
                        group[j].Canonical = false;
                    autoResolvedCount++;
                    Logger.Info(ed, $"dup_latlong autoresolved crossing={displayName}");
                }
                else
                {
                    // Present dialog to choose canonical value
                    using (var dialog = new LatLongDuplicateResolverDialog(group, displayName, i + 1, groups.Count))
                    {
                        if (ModelessDialogRunner.ShowDialog(dialog) != DialogResult.OK)
                            return false;
                    }
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

            Logger.Info(ed, $"dup_latlong summary groups={groups.Count} autoresolved={autoResolvedCount} manual={groups.Count - autoResolvedCount}");

            // ------------------------------------------------------------------
            // Critical part: immediately push chosen LAT/LONG into DWG tables.
            // This prevents the UI from "snapping back" when it refreshes from DWG.
            // ------------------------------------------------------------------
                        // NOTE: We intentionally do not write back to the drawing here.
            // The caller (Apply to Drawing / Update Tables) is responsible for persisting changes.


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
                            // Robust key read
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


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\LayerUtils.cs
/////////////////////////////////////////////////////////////////////

using System;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;

namespace XingManager.Services
{
    /// <summary>
    /// Helper functions for managing layers within the drawing.
    /// </summary>
    public static class LayerUtils
    {
        /// <summary>
        /// Ensures that the supplied layer exists in the database. If the layer is missing it will be
        /// created with default properties compatible with AutoCAD 2014.
        /// </summary>
        public static ObjectId EnsureLayer(Database db, Transaction tr, string layerName)
        {
            if (db == null)
            {
                throw new ArgumentNullException("db");
            }

            if (tr == null)
            {
                throw new ArgumentNullException("tr");
            }

            if (string.IsNullOrWhiteSpace(layerName))
            {
                throw new ArgumentException("Layer name must be supplied", "layerName");
            }

            var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (layerTable.Has(layerName))
            {
                return layerTable[layerName];
            }

            layerTable.UpgradeOpen();
            var layer = new LayerTableRecord
            {
                Name = layerName,
                Color = Color.FromRgb(0, 0, 0)
            };

            var layerId = layerTable.Add(layer);
            tr.AddNewlyCreatedDBObject(layer, true);
            return layerId;
        }
    }
}

/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\LayoutUtils.cs
/////////////////////////////////////////////////////////////////////

using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace XingManager.Services
{
    /// <summary>
    /// Layout and template related helpers.
    /// </summary>
    public class LayoutUtils
    {
        public const string LocationPlaceholder = "_._.1/4 SEC. __, TWP. __, RGE. __, W._M.";
        public const string PlanHeadingBase = "PLAN SHOWING PIPELINE CROSSING(S) WITHIN";
        public const string PlanHeadingAdjacentSuffix = " AND ADJACENT TO";

        private static readonly Regex LocationPlaceholderRegex = new Regex(
            "_\\._\\.1/4\\s+SEC\\.\\s+__,\\s+TWP\\.\\s+__,\\s+RGE\\.\\s+__,\\s+W\\._M\\.",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private const string PlaceholderWs = @"(?:\s+|\\P|\\~)+";

        private static readonly Regex LocationPlaceholderRegexMText = new Regex(
            $"_\\._\\.1/4{PlaceholderWs}SEC\\.{PlaceholderWs}__,{PlaceholderWs}TWP\\.{PlaceholderWs}__,{PlaceholderWs}RGE\\.{PlaceholderWs}__,{PlaceholderWs}W\\._M\\.",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex PlanHeadingRegex = new Regex(
            @"PLAN\s+SHOWING\s+PIPELINE\s+CROSSING\(S\)\s+WITHIN(?:\s+AND\s+ADJACENT\s+TO)?",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public ObjectId CloneLayoutFromTemplate(Document doc, string templatePath, string layoutName, string desiredName, out string actualName)
        {
            if (doc == null)
            {
                throw new ArgumentNullException("doc");
            }

            if (string.IsNullOrWhiteSpace(templatePath))
            {
                throw new ArgumentException("Template path is required", "templatePath");
            }

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException("Template not found", templatePath);
            }

            if (string.IsNullOrWhiteSpace(layoutName))
            {
                throw new ArgumentException("Layout name is required", "layoutName");
            }

            actualName = desiredName ?? layoutName;
            var db = doc.Database;
            ObjectId clonedLayoutId = ObjectId.Null;

            using (var templateDb = new Database(false, true))
            {
                templateDb.ReadDwgFile(templatePath, FileShare.Read, true, string.Empty);
                using (var sourceTr = templateDb.TransactionManager.StartTransaction())
                using (var targetTr = db.TransactionManager.StartTransaction())
                {
                    var layoutDict = (DBDictionary)sourceTr.GetObject(templateDb.LayoutDictionaryId, OpenMode.ForRead);
                    if (!layoutDict.Contains(layoutName))
                    {
                        throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, "Layout '{0}' not found in template", layoutName));
                    }

                    var sourceLayoutId = layoutDict.GetAt(layoutName);
                    var sourceLayout = (Layout)sourceTr.GetObject(sourceLayoutId, OpenMode.ForRead);
                    var sourceBtrId = sourceLayout.BlockTableRecordId;

                    var targetLayoutDict = (DBDictionary)targetTr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                    // Clone the block table record first.
                    var blockIds = new ObjectIdCollection { sourceBtrId };
                    var mapping = new IdMapping();
                    templateDb.WblockCloneObjects(blockIds, db.BlockTableId, mapping, DuplicateRecordCloning.Ignore, false);

                    ObjectId clonedBtrId = ObjectId.Null;
                    foreach (IdPair pair in mapping)
                    {
                        if (pair.Key == sourceBtrId)
                        {
                            clonedBtrId = pair.Value;
                        }
                    }

                    // Now clone the layout entry itself.
                    var layoutIds = new ObjectIdCollection { sourceLayoutId };
                    var layoutMapping = new IdMapping();
                    templateDb.WblockCloneObjects(layoutIds, db.LayoutDictionaryId, layoutMapping, DuplicateRecordCloning.Ignore, false);

                    foreach (IdPair pair in layoutMapping)
                    {
                        if (pair.Key == sourceLayoutId)
                        {
                            clonedLayoutId = pair.Value;
                        }
                    }

                    if (!clonedLayoutId.IsNull)
                    {
                        var layout = (Layout)targetTr.GetObject(clonedLayoutId, OpenMode.ForWrite);
                        if (!clonedBtrId.IsNull)
                        {
                            layout.BlockTableRecordId = clonedBtrId;
                        }

                        actualName = EnsureUniqueLayoutName(targetLayoutDict, layout.LayoutName, desiredName);
                        if (!string.Equals(layout.LayoutName, actualName, StringComparison.OrdinalIgnoreCase))
                        {
                            layout.LayoutName = actualName;
                        }
                    }

                    targetTr.Commit();
                }
            }

            return clonedLayoutId;
        }

        private static string EnsureUniqueLayoutName(DBDictionary layoutDict, string currentName, string desiredName)
        {
            var layoutName = string.IsNullOrWhiteSpace(desiredName) ? currentName : desiredName;
            var uniqueName = layoutName;
            var index = 1;
            while (layoutDict.Contains(uniqueName))
            {
                uniqueName = string.Format(CultureInfo.InvariantCulture, "{0}-{1}", layoutName, index++);
            }

            return uniqueName;
        }

        public void SwitchToLayout(Document doc, string layoutName)
        {
            if (doc == null)
            {
                throw new ArgumentNullException("doc");
            }

            if (string.IsNullOrWhiteSpace(layoutName))
            {
                return;
            }

            var layoutManager = LayoutManager.Current;
            try
            {
                layoutManager.CurrentLayout = layoutName;
            }
            catch (System.Exception ex)
            {
                CommandLogger.Log(doc.Editor, $"Unable to switch layout: {ex.Message}");
            }
        }

        public void ReplacePlaceholderText(Database db, ObjectId layoutId, string replacement)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                bool changed = false;

                foreach (ObjectId entId in btr)
                {
                    var ent = tr.GetObject(entId, OpenMode.ForRead);

                    // ----- Top-level DBTEXT -----
                    if (ent is DBText dbText)
                    {
                        var text = dbText.TextString ?? string.Empty;
                        if (TryReplaceLocationPlaceholder(ref text, replacement))
                        {
                            dbText.UpgradeOpen();
                            dbText.TextString = text;
                            changed = true;
                        }
                        continue;
                    }

                    // ----- Top-level MTEXT -----
                    if (ent is MText mtext)
                    {
                        var raw = mtext.Contents ?? string.Empty;
                        if (TryReplaceLocationPlaceholderInMText(ref raw, replacement))
                        {
                            mtext.UpgradeOpen();
                            mtext.Contents = raw;
                            changed = true;
                            continue;
                        }

                        var plain = mtext.Text ?? mtext.Contents ?? string.Empty;
                        if (IsPlaceholderLoose(plain))
                        {
                            raw = mtext.Contents ?? string.Empty;
                            var prefix = Regex.Match(raw, @"^\s*(?:\\[^;]+;|{[^}]*})*").Value; // keep \H, \f etc.
                            mtext.UpgradeOpen();
                            mtext.Contents = prefix + replacement;
                            changed = true;
                        }
                        continue;
                    }

                    // ----- Attributes on a title block -----
                    if (ent is BlockReference br && br.AttributeCollection != null)
                    {
                        foreach (ObjectId attId in br.AttributeCollection)
                        {
                            if (!(tr.GetObject(attId, OpenMode.ForRead) is AttributeReference attRef)) continue;

                            if (attRef.IsMTextAttribute && attRef.MTextAttribute != null)
                            {
                                var raw = attRef.MTextAttribute.Contents ?? attRef.TextString ?? string.Empty;
                                if (TryReplaceLocationPlaceholderInMText(ref raw, replacement))
                                {
                                    attRef.UpgradeOpen();
                                    attRef.MTextAttribute.Contents = raw;
                                    attRef.TextString = attRef.MTextAttribute.Text;
                                    changed = true;
                                    continue;
                                }

                                var plainMText = attRef.MTextAttribute.Text ?? attRef.TextString ?? string.Empty;
                                if (IsPlaceholderLoose(plainMText))
                                {
                                    var prefix = Regex.Match(attRef.MTextAttribute.Contents ?? string.Empty, @"^\s*(?:\\[^;]+;|{[^}]*})*").Value;
                                    attRef.UpgradeOpen();
                                    attRef.MTextAttribute.Contents = prefix + replacement;
                                    attRef.TextString = attRef.MTextAttribute.Text;
                                    changed = true;
                                }

                                continue;
                            }

                            var text = attRef.TextString ?? string.Empty;
                            if (TryReplaceLocationPlaceholder(ref text, replacement))
                            {
                                attRef.UpgradeOpen();
                                attRef.TextString = text;
                                changed = true;
                            }
                        }
                    }
                }

                tr.Commit();
                if (changed)
                {
                    try { Application.UpdateScreen(); } catch { }
                }
            }
        }

        public void UpdatePlanHeadingText(Database db, ObjectId layoutId, bool includeAdjacent)
        {
            if (db == null || layoutId.IsNull) return;

            var replacement = includeAdjacent
                ? PlanHeadingBase + PlanHeadingAdjacentSuffix
                : PlanHeadingBase;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

                foreach (ObjectId entId in btr)
                {
                    // DBTEXT: safe to replace via TextString (no inline formatting)
                    if (tr.GetObject(entId, OpenMode.ForRead) is DBText dbText)
                    {
                        var txt = dbText.TextString ?? string.Empty;
                        if (TryUpdateHeading(ref txt, replacement))
                        {
                            dbText.UpgradeOpen();
                            dbText.TextString = txt;
                        }
                        continue;
                    }

                    // MTEXT: operate on Contents (raw, with \H, \f etc.) so we don't lose formatting
                    if (tr.GetObject(entId, OpenMode.ForRead) is MText mtext)
                    {
                        var raw = mtext.Contents ?? string.Empty;
                        if (TryUpdateHeadingInMText(ref raw, replacement))
                        {
                            mtext.UpgradeOpen();
                            mtext.Contents = raw;   // formatting preserved
                        }
                    }
                }

                tr.Commit();
            }
        }


        // Variant that also treats MTEXT nonâ€‘breaking spaces (\~) as whitespace
        private static readonly string Ws = @"(?:\s+|\\~)+";
        private static readonly Regex PlanHeadingRegexMText = new Regex(
            $"PLAN{Ws}SHOWING{Ws}PIPELINE{Ws}CROSSING\\(S\\){Ws}WITHIN(?:{Ws}AND{Ws}ADJACENT{Ws}TO)?",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static bool TryUpdateHeadingInMText(ref string contents, string replacement)
        {
            if (string.IsNullOrEmpty(contents)) return false;

            // 1) Try straight replace in the raw contents (keeps leading \H, \f, etc.)
            var updated = PlanHeadingRegex.Replace(contents, replacement);

            // 2) If not found, try the MTEXTâ€‘aware variant that understands \~
            if (updated == contents)
                updated = PlanHeadingRegexMText.Replace(contents, replacement);

            if (updated == contents) return false;

            contents = updated;
            return true;
        }

        private static bool TryUpdateHeading(ref string text, string replacement)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            var match = PlanHeadingRegex.Match(text);
            if (!match.Success)
            {
                return false;
            }

            var updated = PlanHeadingRegex.Replace(text, replacement);
            if (string.Equals(updated, text, StringComparison.Ordinal))
            {
                return false;
            }

            text = updated;
            return true;
        }

        private static bool IsPlaceholderLoose(string value)
        {
            var s = (value ?? string.Empty)
                .Replace('\u00A0', ' ')        // NBSP -> space
                .Trim();
            s = Regex.Replace(s, @"\s+", " "); // collapse whitespace
            return string.Equals(s, LocationPlaceholder, StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryReplaceLocationPlaceholder(ref string text, string replacement)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            var updated = LocationPlaceholderRegex.Replace(text, replacement);
            if (!string.Equals(updated, text, StringComparison.Ordinal))
            {
                text = updated;
                return true;
            }

            if (IsPlaceholderLoose(text))
            {
                text = replacement;
                return true;
            }

            return false;
        }

        private static bool TryReplaceLocationPlaceholderInMText(ref string contents, string replacement)
        {
            if (string.IsNullOrEmpty(contents)) return false;

            // If the normalized raw contents exactly equals the placeholder,
            // replace the WHOLE thing while preserving leading/trailing formatting.
            if (RawMTextEqualsPlaceholder(contents))
            {
                var prefix = LeadingFmt.Match(contents).Value;
                var suffix = TrailingFmt.Match(contents).Value;
                contents = prefix + replacement + suffix;
                return true;
            }

            // Fallback: try a tolerant raw replace that also recognizes stacked fraction \S1/4;
            const string WS = @"(?:\s+|\\P|\\~)+";
            const string F = @"(?:\\[A-Za-z][^;]*;|{[^}]*})*";      // any inline format tokens
            const string FR = @"(?:1/4|\\S1[/#]4;)";                 // literal 1/4 or stacked fraction

            var placeholderRaw = new Regex(
                $"{F}_\\.{F}_\\.{F}{FR}{WS}SEC\\.{WS}__,{WS}TWP\\.{WS}__,{WS}RGE\\.{WS}__,{WS}W\\.{F}_M\\.{F}",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

            var updated = placeholderRaw.Replace(contents, m =>
            {
                // Try to keep whatever leading/trailing format was on the match
                var lead = LeadingFmt.Match(m.Value).Value;
                var tail = TrailingFmt.Match(m.Value).Value;
                return lead + replacement + tail;
            });

            if (!string.Equals(updated, contents, StringComparison.Ordinal))
            {
                contents = updated;
                return true;
            }

            return false;
        }
        // Strip inline MTEXT codes into comparable plain text.
        // - converts \S1/4; or \S1#4; to "1/4"
        // - drops \H, \f, \A, \Q, \W, \L, \O, etc.
        // - treats \P (newline) and \~ (nbsp) as spaces
        private static string NormalizeMTextForComparison(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return string.Empty;

            var s = raw;

            // 1) Expand stacked fractions: \S1/4; or \S1#4;
            s = Regex.Replace(s, @"\\S([^;]+);", m =>
            {
                var frac = m.Groups[1].Value.Replace('#', '/');
                return frac;
            }, RegexOptions.IgnoreCase);

            // 2) Replace \P (paragraph) and \~ (NBSP) with space
            s = Regex.Replace(s, @"\\P|\\~", " ", RegexOptions.IgnoreCase);

            // 3) Drop all other inline formatting codes like \H2.5x;, \fArial|b0|i0;, \A1;, etc.
            s = Regex.Replace(s, @"\\[A-Za-z][^;]*;", string.Empty);

            // 4) Drop braces used to group overrides
            s = s.Replace("{", string.Empty).Replace("}", string.Empty);

            // 5) Normalize whitespace and NBSP
            s = s.Replace('\u00A0', ' ');
            s = Regex.Replace(s, @"\s+", " ").Trim();

            return s;
        }

        private static string NormalizeForComparison(string value)
        {
            var s = (value ?? string.Empty).Replace('\u00A0', ' ');
            s = Regex.Replace(s, @"\s+", " ").Trim();
            return s;
        }

        // Returns true if raw MTEXT contents equals the placeholder once normalized.
        private static bool RawMTextEqualsPlaceholder(string rawContents)
        {
            var left = NormalizeMTextForComparison(rawContents);
            var right = NormalizeForComparison(LocationPlaceholder);
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }

        // Capture leading/trailing formatting runs so we can preserve them.
        private static readonly Regex LeadingFmt = new Regex(@"^\s*(?:\\[A-Za-z][^;]*;|{[^}]*})*", RegexOptions.Compiled);
        private static readonly Regex TrailingFmt = new Regex(@"(?:\\[A-Za-z][^;]*;|{[^}]*})*\s*$", RegexOptions.Compiled);

        public bool TryFormatMeridianLocation(string raw, out string formatted)
        {
            formatted = raw;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return false;
            }

            var pattern = new Regex("^\\s*(?<q>(?:n\\.?\\s*[ew]|s\\.?\\s*[ew]))\\.?\\s*(?<sec>\\d+)\\s*-\\s*(?<twp>\\d+)\\s*-\\s*(?<rge>\\d+)\\s*-\\s*(?<meridian>\\d+)\\s*$", RegexOptions.IgnoreCase);
            var match = pattern.Match(raw);
            if (!match.Success)
            {
                return false;
            }

            var quarter = NormalizeQuarter(match.Groups["q"].Value);
            var sec = match.Groups["sec"].Value;
            var twp = match.Groups["twp"].Value;
            var rge = match.Groups["rge"].Value;
            var meridian = match.Groups["meridian"].Value;

            formatted = string.Format(CultureInfo.InvariantCulture, "{0}1/4 SEC. {1}, TWP. {2}, RGE. {3}, W.{4}M.", quarter, sec, twp, rge, meridian);
            return true;
        }

        private static string NormalizeQuarter(string value)
        {
            var cleaned = new string((value ?? string.Empty).Where(char.IsLetter).ToArray()).ToUpperInvariant();
            switch (cleaned)
            {
                case "NE":
                    return "N.E.";
                case "NW":
                    return "N.W.";
                case "SE":
                    return "S.E.";
                case "SW":
                    return "S.W.";
                default:
                    return cleaned;
            }
        }
    }
}

/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\Logger.cs
/////////////////////////////////////////////////////////////////////

using System;
using System.Diagnostics;
using Autodesk.AutoCAD.EditorInput;

namespace XingManager.Services
{
    public static class Logger
    {
        public enum Level { Debug = 0, Info = 1, Warn = 2, Error = 3 }

        static Logger()
        {
            var env = (Environment.GetEnvironmentVariable("XING_LOG_LEVEL") ?? string.Empty)
                .Trim()
                .ToUpperInvariant();

            if (env == "DEBUG")
            {
                CurrentLevel = Level.Debug;
            }
            else if (env == "WARN")
            {
                CurrentLevel = Level.Warn;
            }
            else if (env == "ERROR")
            {
                CurrentLevel = Level.Error;
            }
            else
            {
                CurrentLevel = Level.Info;
            }
        }

        public static Level CurrentLevel { get; set; }

        public static void Debug(Editor ed, string msg)
        {
            if (CurrentLevel <= Level.Debug) Write(ed, "DEBUG", msg);
        }

        public static void Info(Editor ed, string msg)
        {
            if (CurrentLevel <= Level.Info) Write(ed, "INFO", msg);
        }

        public static void Warn(Editor ed, string msg)
        {
            if (CurrentLevel <= Level.Warn) Write(ed, "WARN", msg);
        }

        public static void Error(Editor ed, string msg)
        {
            Write(ed, "ERROR", msg);
        }

        public static IDisposable Scope(Editor ed, string name, string kv = null)
        {
            Write(ed, "INFO", $"start {name}{(string.IsNullOrEmpty(kv) ? string.Empty : " " + kv)}");
            var sw = Stopwatch.StartNew();
            return new ScopeImpl(() =>
            {
                sw.Stop();
                Write(ed, "INFO", $"end   {name} elapsed_ms={sw.ElapsedMilliseconds}{(string.IsNullOrEmpty(kv) ? string.Empty : " " + kv)}");
            });
        }

        private sealed class ScopeImpl : IDisposable
        {
            private readonly Action _onDispose;

            public ScopeImpl(Action onDispose)
            {
                _onDispose = onDispose;
            }

            public void Dispose()
            {
                try
                {
                    _onDispose?.Invoke();
                }
                catch
                {
                }
            }
        }

        private static void Write(Editor ed, string level, string msg)
        {
            if (string.IsNullOrWhiteSpace(msg))
            {
                return;
            }

            try
            {
                ed?.WriteMessage($"\n[CrossingManager][{level}] {msg}");
            }
            catch
            {
            }
        }
    }
}

/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\ModelessDialogRunner.cs
/////////////////////////////////////////////////////////////////////

using System;
using System.Threading;
using System.Windows.Forms;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace XingManager.Services
{
    /// <summary>
    /// Helper that shows a WinForms dialog modelessly while waiting synchronously
    /// for the user to close it. This allows interaction with the DWG (panning, zooming, etc.)
    /// while the dialog is visible.
    /// </summary>
    internal static class ModelessDialogRunner
    {
        private sealed class WindowHandleWrapper : IWin32Window
        {
            private readonly IntPtr _handle;

            public WindowHandleWrapper(IntPtr handle)
            {
                _handle = handle;
            }

            public IntPtr Handle => _handle;
        }

        public static DialogResult ShowDialog(Form dialog)
        {
            if (dialog == null)
                throw new ArgumentNullException(nameof(dialog));

            var completed = false;
            var result = DialogResult.None;

            void OnClosed(object sender, FormClosedEventArgs args)
            {
                dialog.FormClosed -= OnClosed;
                completed = true;
                result = dialog.DialogResult;
            }

            dialog.FormClosed += OnClosed;

            var mainWindow = AcadApp.MainWindow;
            if (mainWindow != null)
            {
                var owner = new WindowHandleWrapper(mainWindow.Handle);
                dialog.Show(owner);
            }
            else
            {
                dialog.Show();
            }

            while (!completed)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(25);
            }

            return result;
        }
    }
}

/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\Serde.cs
/////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using XingManager.Models;

namespace XingManager.Services
{
    /// <summary>
    /// Handles CSV import and export for crossing data.
    /// </summary>
    public class Serde
    {
        private static readonly string[] Header = { "CROSSING", "OWNER", "DESCRIPTION", "LOCATION", "DWG_REF", "LAT", "LONG", "ZONE" };

        public void Export(string path, IEnumerable<CrossingRecord> records)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Path must be provided", "path");
            }

            var list = records == null
                ? new List<CrossingRecord>()
                : records.OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing)).ToList();

            using (var writer = new StreamWriter(path, false, Encoding.UTF8))
            {
                writer.WriteLine(string.Join(",", Header));
                foreach (var record in list)
                {
                    var values = new[]
                    {
                        Escape(record.Crossing),
                        Escape(record.Owner),
                        Escape(record.Description),
                        Escape(record.Location),
                        Escape(record.DwgRef),
                        Escape(record.Lat),
                        Escape(record.Long),
                        Escape(record.Zone)
                    };

                    writer.WriteLine(string.Join(",", values));
                }
            }
        }

        public List<CrossingRecord> Import(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Path must be provided", "path");
            }

            if (!File.Exists(path))
            {
                throw new FileNotFoundException("CSV file not found", path);
            }

            var result = new List<CrossingRecord>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var reader = new StreamReader(path, Encoding.UTF8))
            {
                string line;
                var isFirst = true;
                while ((line = reader.ReadLine()) != null)
                {
                    if (isFirst)
                    {
                        isFirst = false;
                        continue; // Skip header.
                    }

                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    var fields = ParseCsvLine(line);
                    if (fields.Count < Header.Length)
                    {
                        fields.AddRange(Enumerable.Repeat(string.Empty, Header.Length - fields.Count));
                    }

                    var record = new CrossingRecord
                    {
                        Crossing = fields[0],
                        Owner = fields[1],
                        Description = fields[2],
                        Location = fields[3],
                        DwgRef = fields[4],
                        Lat = fields[5],
                        Long = fields[6],
                        Zone = fields.Count > 7 ? fields[7] : string.Empty
                    };

                    var key = record.CrossingKey;
                    if (string.IsNullOrEmpty(key))
                    {
                        throw new InvalidDataException("CROSSING value is required for all rows");
                    }

                    if (!seen.Add(key))
                    {
                        throw new InvalidDataException(string.Format(CultureInfo.InvariantCulture, "Duplicate CROSSING '{0}' detected in CSV", record.Crossing));
                    }

                    result.Add(record);
                }
            }

            return result;
        }

        private static string Escape(string value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
            {
                return string.Format(CultureInfo.InvariantCulture, "\"{0}\"", value.Replace("\"", "\"\""));
            }

            return value;
        }

        private static List<string> ParseCsvLine(string line)
        {
            var values = new List<string>();
            var sb = new StringBuilder();
            var inQuotes = false;

            for (var i = 0; i < line.Length; i++)
            {
                var c = line[i];
                if (c == '\"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '\"')
                    {
                        sb.Append('\"');
                        i++;
                        continue;
                    }

                    inQuotes = !inQuotes;
                    continue;
                }

                if (c == ',' && !inQuotes)
                {
                    values.Add(sb.ToString());
                    sb.Length = 0;
                    continue;
                }

                sb.Append(c);
            }

            values.Add(sb.ToString());
            return values;
        }
    }
}

/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\TableCellProbe.cs
/////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace XingManager.Services
{
    public static class TableCellProbe
    {
        public static string TryGetCellBlockAttr(Table t, int row, int col, string tag)
        {
            if (t == null || row < 0 || col < 0 || string.IsNullOrWhiteSpace(tag)) return string.Empty;

            // 1) (row, col, tag, Ã¢â‚¬Â¦)
            var v = TryCallGetBlockAttr(t, row, col, tag);
            if (!string.IsNullOrWhiteSpace(v)) return v.Trim();

            // 2) iterate contents and try indexed overload
            var cell = SafeGetCell(t, row, col);
            var contents = GetContents(cell);
            int idx = 0;
            foreach (var _ in contents)
            {
                v = TryCallGetBlockAttrIndexed(t, row, col, idx, tag);
                if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
                idx++;
            }

            // 3) discover tags from the cellÃ¢â‚¬â„¢s block definition and try those
            foreach (var discovered in EnumerateCellBlockTags(t, row, col))
            {
                v = TryCallGetBlockAttr(t, row, col, discovered);
                if (!string.IsNullOrWhiteSpace(v)) return v.Trim();

                idx = 0;
                foreach (var __ in contents)
                {
                    v = TryCallGetBlockAttrIndexed(t, row, col, idx, discovered);
                    if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
                    idx++;
                }
            }

            DebugLog($"table_cell_probe row={row} col={col} tag={tag} status=no_match");
            return string.Empty;
        }

        // -------- internals --------
        private static Cell SafeGetCell(Table t, int row, int col)
        {
            try { return t.Cells[row, col]; } catch { return null; }
        }

        private static IEnumerable GetContents(Cell cell)
        {
            if (cell == null) yield break;
            var p = cell.GetType().GetProperty("Contents", BindingFlags.Public | BindingFlags.Instance);
            var contents = p?.GetValue(cell, null) as IEnumerable;
            if (contents == null) yield break;
            foreach (var c in contents) yield return c;
        }

        private static string TryCallGetBlockAttr(Table t, int row, int col, string tag)
        {
            const string name = "GetBlockAttributeValue";
            foreach (var mi in t.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance))
            {
                if (!string.Equals(mi.Name, name, StringComparison.Ordinal)) continue;
                var p = mi.GetParameters();
                if (p.Length < 3) continue; // need at least (row, col, tag)

                // expect (row, col, string tag, Ã¢â‚¬Â¦)
                if (typeof(string).IsAssignableFrom(p[2].ParameterType))
                {
                    var args = new object[p.Length];
                    if (!TryConvert(row, p[0], out args[0]) ||
                        !TryConvert(col, p[1], out args[1]) ||
                        !TryConvert(tag, p[2], out args[2])) continue;

                    for (int i = 3; i < p.Length; i++) args[i] = p[i].IsOptional ? Type.Missing : null;

                    try { return Convert.ToString(mi.Invoke(t, args)); } catch (Exception ex) { DebugLog($"table_cell_probe reflection_fail mode=direct method={mi.Name} err={ex.Message}"); }
                }
            }
            DebugLog($"table_cell_probe reflection_miss mode=direct row={row} col={col} tag={tag}");
            return string.Empty;
        }

        private static string TryCallGetBlockAttrIndexed(Table t, int row, int col, int contentIndex, string tag)
        {
            const string name = "GetBlockAttributeValue";
            foreach (var mi in t.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance))
            {
                if (!string.Equals(mi.Name, name, StringComparison.Ordinal)) continue;
                var p = mi.GetParameters();
                if (p.Length < 4) continue; // need (row, col, contentIndex, tag)

                // expect (row, col, int contentIndex, string tag, Ã¢â‚¬Â¦)
                if (typeof(string).IsAssignableFrom(p[3].ParameterType))
                {
                    var args = new object[p.Length];
                    if (!TryConvert(row,   p[0], out args[0]) ||
                        !TryConvert(col,   p[1], out args[1]) ||
                        !TryConvert(contentIndex, p[2], out args[2]) ||
                        !TryConvert(tag,   p[3], out args[3])) continue;

                    for (int i = 4; i < p.Length; i++) args[i] = p[i].IsOptional ? Type.Missing : null;

                    try { return Convert.ToString(mi.Invoke(t, args)); } catch (Exception ex) { DebugLog($"table_cell_probe reflection_fail mode=indexed method={mi.Name} err={ex.Message}"); }
                }
            }
            DebugLog($"table_cell_probe reflection_miss mode=indexed row={row} col={col} idx={contentIndex} tag={tag}");
            return string.Empty;
        }

        private static bool TryConvert(object value, ParameterInfo p, out object converted)
        {
            var target = Nullable.GetUnderlyingType(p.ParameterType) ?? p.ParameterType;
            try
            {
                if (value == null) { converted = target.IsValueType ? Activator.CreateInstance(target) : null; return true; }
                if (target.IsInstanceOfType(value)) { converted = value; return true; }
                converted = Convert.ChangeType(value, target, System.Globalization.CultureInfo.InvariantCulture);
                return true;
            }
            catch { converted = null; return false; }
        }

        private static IEnumerable<string> EnumerateCellBlockTags(Table t, int row, int col)
        {
            var cell = SafeGetCell(t, row, col);
            var contents = GetContents(cell);
            var tr = t.Database?.TransactionManager?.TopTransaction as Transaction;
            if (tr == null) yield break;

            foreach (var content in contents)
            {
                var ctProp = content.GetType().GetProperty("ContentTypes", BindingFlags.Public | BindingFlags.Instance);
                var types = ctProp?.GetValue(content, null)?.ToString() ?? string.Empty;
                if (types.IndexOf("Block", StringComparison.OrdinalIgnoreCase) < 0) continue;

                var btrProp = content.GetType().GetProperty("BlockTableRecordId", BindingFlags.Public | BindingFlags.Instance);
                if (!(btrProp?.GetValue(content, null) is ObjectId btrId) || btrId.IsNull || !btrId.IsValid) continue;

                BlockTableRecord btr = null;
                try { btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord; } catch { }
                if (btr == null) continue;

                foreach (ObjectId eid in btr)
                {
                    AttributeDefinition ad = null;
                    try { ad = tr.GetObject(eid, OpenMode.ForRead) as AttributeDefinition; } catch { }
                    if (ad != null && !string.IsNullOrWhiteSpace(ad.Tag)) yield return ad.Tag.Trim();
                }
            }
        }

        private static void DebugLog(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
                return;

            var ed = Application.DocumentManager?.MdiActiveDocument?.Editor;
            Logger.Debug(ed, message);
        }
    }
}

/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\TableFactory.cs
/////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using XingManager.Models;

namespace XingManager.Services
{
    /// <summary>
    /// Handles the creation of new tables using AutoCAD 2014 compatible APIs.
    /// </summary>
    public class TableFactory
    {
        public const string TableStyleName = "Induction bend";
        public const string TableTypeXrecordKey = "XING_TABLE_TYPE";
        public const string LayerName = "CG-NOTES";

        public TableFactory()
        {
        }

        private static void SetCellText(Table t, int row, int col, string text, double height, ObjectId textStyleId, CellAlignment align = CellAlignment.MiddleLeft)
        {
            var cell = t.Cells[row, col];
            cell.Contents.Clear();
            cell.TextString = text ?? string.Empty;
            cell.TextHeight = height;
            cell.TextStyleId = textStyleId;
            cell.Alignment = align;
        }

        public ObjectId EnsureTableStyle(Database db, Transaction tr)
        {
            if (db == null) throw new ArgumentNullException("db");
            if (tr == null) throw new ArgumentNullException("tr");

            var styleDict = (DBDictionary)tr.GetObject(db.TableStyleDictionaryId, OpenMode.ForRead);
            if (styleDict.Contains(TableStyleName))
                return styleDict.GetAt(TableStyleName);

            styleDict.UpgradeOpen();
            var tableStyle = new TableStyle
            {
                Name = TableStyleName
            };

            // No SetTextHeight/SetTextStyle calls here; we format per-cell for 2014+ compatibility.
            var id = styleDict.SetAt(TableStyleName, tableStyle);
            tr.AddNewlyCreatedDBObject(tableStyle, true);
            return id;
        }

        public Table CreateMainCrossingTable(Database db, Transaction tr, IEnumerable<CrossingRecord> records)
        {
            var recordList = PrepareRecordList(records);

            var ed = Application.DocumentManager?.MdiActiveDocument?.Editor;

            var table = new Table
            {
                TableStyle = EnsureTableStyle(db, tr),
                LayerId = LayerUtils.EnsureLayer(db, tr, LayerName)
            };

            var rows = recordList.Count + 1;
            // NumRows/NumColumns are obsolete in newer AutoCAD APIs.
            // Use SetSize(rows, cols) to stay forward-compatible.
            table.SetSize(rows, 5);

            table.SetRowHeight(25.0);
            table.Columns[0].Width = 43.5;
            table.Columns[1].Width = 144.5;
            table.Columns[2].Width = 393.5;
            table.Columns[3].Width = 200.0;
            table.Columns[4].Width = 120.0;

            var textStyleId = db.Textstyle;
            const double textHeight = 10.0;
            var headers = new[] { "XING", "OWNER", "DESCRIPTION", "LOCATION", "DWG_REF" };
            for (var col = 0; col < headers.Length; col++)
            {
                SetCellText(table, 0, col, headers[col], textHeight, textStyleId);
            }

            for (var row = 0; row < recordList.Count; row++)
            {
                var record = recordList[row];
                var rowIndex = row + 1;

                SetCellText(table, rowIndex, 0, record.Crossing, textHeight, textStyleId);
                SetCellText(table, rowIndex, 1, record.Owner, textHeight, textStyleId);
                SetCellText(table, rowIndex, 2, record.Description, textHeight, textStyleId);
                SetCellText(table, rowIndex, 3, record.Location, textHeight, textStyleId);
                SetCellText(table, rowIndex, 4, record.DwgRef, textHeight, textStyleId);
            }

            table.GenerateLayout();
            Logger.Info(ed, $"create_table kind=MAIN rows={recordList.Count}");
            return table;
        }

        public Table CreateCrossingPageTable(Database db, Transaction tr, string dwgRef, IEnumerable<CrossingRecord> records)
        {
            var filtered = PrepareRecordList(records)
                .Where(r => string.Equals((r.DwgRef ?? string.Empty).Trim(), (dwgRef ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase))
                .ToList();

            var ed = Application.DocumentManager?.MdiActiveDocument?.Editor;

            var table = new Table
            {
                TableStyle = EnsureTableStyle(db, tr),
                LayerId = LayerUtils.EnsureLayer(db, tr, LayerName)
            };

            var rows = filtered.Count + 1;
            // NumRows/NumColumns are obsolete in newer AutoCAD APIs.
            // Use SetSize(rows, cols) to stay forward-compatible.
            table.SetSize(Math.Max(1, rows), 3);
            table.SetRowHeight(25.0);
            table.Columns[0].Width = 43.5;
            table.Columns[1].Width = 144.5;
            table.Columns[2].Width = 393.5;

            var textStyleId = db.Textstyle;
            const double textHeight = 10.0;
            SetCellText(table, 0, 0, "XING", textHeight, textStyleId);
            SetCellText(table, 0, 1, "OWNER", textHeight, textStyleId);
            SetCellText(table, 0, 2, "DESCRIPTION", textHeight, textStyleId);

            for (var row = 0; row < filtered.Count; row++)
            {
                var record = filtered[row];
                var rowIndex = row + 1;
                SetCellText(table, rowIndex, 0, record.Crossing, textHeight, textStyleId);
                SetCellText(table, rowIndex, 1, record.Owner, textHeight, textStyleId);
                SetCellText(table, rowIndex, 2, record.Description, textHeight, textStyleId);
            }

            table.GenerateLayout();
            Logger.Info(ed, $"create_table kind=PAGE rows={filtered.Count} dwg={dwgRef}");
            return table;
        }

        public void TagTable(Transaction tr, Table table, string tableType)
        {
            if (table == null)
            {
                throw new ArgumentNullException("table");
            }

            if (tr == null)
            {
                throw new ArgumentNullException("tr");
            }

            if (string.IsNullOrEmpty(tableType))
            {
                return;
            }

            if (table.ExtensionDictionary.IsNull)
            {
                table.UpgradeOpen();
                table.CreateExtensionDictionary();
            }

            var dict = (DBDictionary)tr.GetObject(table.ExtensionDictionary, OpenMode.ForWrite);
            Xrecord xrec;
            if (dict.Contains(TableTypeXrecordKey))
            {
                xrec = (Xrecord)tr.GetObject(dict.GetAt(TableTypeXrecordKey), OpenMode.ForWrite);
            }
            else
            {
                xrec = new Xrecord();
                dict.SetAt(TableTypeXrecordKey, xrec);
                tr.AddNewlyCreatedDBObject(xrec, true);
            }

            xrec.Data = new ResultBuffer(new TypedValue((int)DxfCode.Text, tableType));
            var ed = Application.DocumentManager?.MdiActiveDocument?.Editor;
            Logger.Info(ed, $"tag_table handle={table.ObjectId.Handle} type={tableType}");
        }

        private static List<CrossingRecord> PrepareRecordList(IEnumerable<CrossingRecord> records)
        {
            if (records == null)
            {
                return new List<CrossingRecord>();
            }

            return records
                .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();
        }
    }
}

/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\TableMatcher.cs
/////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using XingManager.Models;

namespace XingManager.Services
{
    /// <summary>
    /// Reads a selected crossing table and pushes those values into XING2 blocks.
    /// IMPORTANT: Match Table does NOT renumber or change X# values. It matches by existing X# and updates the other fields.
    /// Column A is a block cell; we read the CROSSING value from the cell's block attributes.
    /// Adjacent cells provide OWNER / DESCRIPTION / LOCATION / DWG_REF.
    /// </summary>
    public class TableMatcher
    {
        // --- Tag groups used when reading/writing block attributes ---
        private static readonly ISet<string> CrossingAttributeTags = new HashSet<string>(new[]
        {
            "CROSSING","XING","X_NO","XNUM","XNUMBER","NUMBER","INDEX","NO","LABEL"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> OwnerAttributeTags = new HashSet<string>(new[]
        {
            "OWNER","OWN","COMPANY"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> DescriptionAttributeTags = new HashSet<string>(new[]
        {
            "DESCRIPTION","DESC"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> LocationAttributeTags = new HashSet<string>(new[]
        {
            "LOCATION","LOC"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> DwgRefAttributeTags = new HashSet<string>(new[]
        {
            "DWG_REF","DWGREF","DWGREFNO","DWGREFNUMBER"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> LatAttributeTags = new HashSet<string>(new[]
        {
            "LAT","LATITUDE"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> LongAttributeTags = new HashSet<string>(new[]
        {
            "LONG","LONGITUDE"
        }, StringComparer.OrdinalIgnoreCase);

        private static readonly ISet<string> ZoneAttributeTags = new HashSet<string>(new[]
        {
            "ZONE","ZONE_LABEL"
        }, StringComparer.OrdinalIgnoreCase);

        // --- simple normalizer for composite matching keys ---
        private static string N(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            s = s.Trim();
            while (s.Contains("  ")) s = s.Replace("  ", " ");
            return s;
        }
        private static string CompositeKeyMain(string owner, string desc, string loc, string dwg)
            => string.Join("|", N(owner), N(desc), N(loc), N(dwg));
        private static string CompositeKeyPage(string owner, string desc)
            => string.Join("|", N(owner), N(desc));

        [CommandMethod("XING_MATCH_TABLE")]
        public void MatchTable()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            using (Logger.Scope(ed, "match_table"))
            {
                try
                {
                    var options = new PromptEntityOptions("\nSelect a crossing table to match fields (does NOT renumber X#):")
                    {
                        AllowNone = false
                    };
                    options.SetRejectMessage("\nEntity must be a table.");
                    options.AddAllowedClass(typeof(Table), true);

                    var selection = ed.GetEntity(options);
                    if (selection.Status != PromptStatus.OK)
                    {
                        Logger.Info(ed, "match_table status=cancelled");
                        return;
                    }

                using (doc.LockDocument())
                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    var table = tr.GetObject(selection.ObjectId, OpenMode.ForRead) as Table;
                    if (table == null)
                    {
                        Logger.Warn(ed, "match_table status=aborted reason=not_table");
                        return;
                    }

                    // Determine table type with your existing logic.
                    var tableType = DetectTableType(table, tr);
                    if (tableType == TableSync.XingTableType.Unknown)
                    {
                        Logger.Warn(ed, "match_table status=aborted reason=unknown_table_type");
                        return;
                    }
                    Logger.Info(ed, $"table_detected type={tableType}");

                    // Build lookup dictionaries FROM the selected table (this is the "source of truth")
                    HashSet<string> duplicateKeys;
                    Dictionary<string, CrossingRecord> byKey, byComposite;
                    Dictionary<string, string> byCompositeXKey;
                    BuildIndexesFromTable(
                        table,
                        tableType,
                        ed,
                        out byKey,
                        out byComposite,
                        out byCompositeXKey,
                        out duplicateKeys);

                    Logger.Info(ed, $"indexed byKey={byKey.Count} byComposite={byComposite.Count} dupes={duplicateKeys.Count}");
                    Logger.Info(ed, "NOTE: Match Table updates block fields using the SAME X#. It does NOT renumber or change X# values.");

                    // --------------------------------------------------------------------------------------
                    // When matching against a Main/Page table, enrich the in-memory records with LAT/LONG
                    // values from any LAT/LONG tables in the drawing **and** from existing XING blocks.
                    // Without this step, records created from Main/Page tables do not populate the
                    // Lat/Long fields and therefore will not update block extension dictionaries.
                    //
                    // First, harvest LAT/LONG values from all other LAT/LONG tables in the drawing.
                    // We walk through every table in every layout, identify those that are LAT/LONG tables,
                    // build temporary indexes from them, and merge their Lat/Long/Zone values into the
                    // primary byKey dictionary. Only missing fields are populated; if the selected table
                    // already contains LAT/LONG values, they are preserved.
                    if (tableType != TableSync.XingTableType.LatLong)
                    {
                        try
                        {
                            // Collect LAT/LONG information from other tables
                            var latLongByKey = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
                            var tableEnriched = 0;
                            var dbForLat = doc.Database;
                            var btable = (BlockTable)tr.GetObject(dbForLat.BlockTableId, OpenMode.ForRead);
                            foreach (ObjectId btrId2 in btable)
                            {
                                var btr2 = (BlockTableRecord)tr.GetObject(btrId2, OpenMode.ForRead);
                                if (!btr2.IsLayout) continue;
                                foreach (ObjectId entId2 in btr2)
                                {
                                    var tObj = tr.GetObject(entId2, OpenMode.ForRead) as Table;
                                    if (tObj == null) continue;
                                    // skip the selected table
                                    if (tObj.ObjectId == table.ObjectId) continue;

                                    TableSync.XingTableType tType = TableSync.XingTableType.Unknown;
                                    try
                                    {
                                        tType = DetectTableType(tObj, tr);
                                    }
                                    catch { }
                                    if (tType != TableSync.XingTableType.LatLong) continue;

                                    // Build indexes from this lat-long table
                                    BuildIndexesFromTable(
                                        tObj,
                                        tType,
                                        ed,
                                        out var llByKey,
                                        out var llByComposite,
                                        out var llByCompositeX,
                                        out var llDupes,
                                        logDuplicates: false);
                                    foreach (var kv in llByKey)
                                    {
                                        if (!latLongByKey.TryGetValue(kv.Key, out var merged))
                                        {
                                            latLongByKey[kv.Key] = kv.Value;
                                        }
                                        else
                                        {
                                            // combine lat/long/zone if missing
                                            if (string.IsNullOrWhiteSpace(merged.Lat) && !string.IsNullOrWhiteSpace(kv.Value.Lat))
                                                merged.Lat = kv.Value.Lat;
                                            if (string.IsNullOrWhiteSpace(merged.Long) && !string.IsNullOrWhiteSpace(kv.Value.Long))
                                                merged.Long = kv.Value.Long;
                                            if (string.IsNullOrWhiteSpace(merged.Zone) && !string.IsNullOrWhiteSpace(kv.Value.Zone))
                                                merged.Zone = kv.Value.Zone;
                                        }
                                    }
                                }
                            }
                            // Merge collected LAT/LONG values into byKey
                            foreach (var kv in latLongByKey)
                            {
                                if (byKey.TryGetValue(kv.Key, out var recTarget))
                                {
                                    var llRec = kv.Value;
                                    var changedFromTable = false;
                                    if (string.IsNullOrWhiteSpace(recTarget.Lat) && !string.IsNullOrWhiteSpace(llRec.Lat))
                                    {
                                        recTarget.Lat = llRec.Lat;
                                        changedFromTable = true;
                                    }
                                    if (string.IsNullOrWhiteSpace(recTarget.Long) && !string.IsNullOrWhiteSpace(llRec.Long))
                                    {
                                        recTarget.Long = llRec.Long;
                                        changedFromTable = true;
                                    }
                                    if (string.IsNullOrWhiteSpace(recTarget.Zone) && !string.IsNullOrWhiteSpace(llRec.Zone))
                                    {
                                        recTarget.Zone = llRec.Zone;
                                        changedFromTable = true;
                                    }
                                    if (changedFromTable)
                                        tableEnriched++;
                                }
                            }
                            Logger.Info(ed, $"enriched_from source=latlong_tables keys={tableEnriched}");
                        }
                        catch (System.Exception exTables)
                        {
                            // Best effort: do not abort if we cannot enrich from tables
                            Logger.Warn(ed, $"enrich_latlong_from_tables_failed err={exTables.Message}");
                        }

                        try
                        {
                            // Secondly, harvest LAT/LONG values from existing XING blocks via repository scan
                            var scanResult = new XingRepository(doc).ScanCrossings();
                            var blockEnriched = 0;
                            foreach (var record in scanResult.Records ?? new List<CrossingRecord>())
                            {
                                // Normalize the key using the same normalizer as BuildIndexesFromTable
                                var normalized = TableSync.NormalizeKeyForLookup(record.Crossing);
                                if (string.IsNullOrEmpty(normalized))
                                    continue;

                                if (byKey.TryGetValue(normalized, out var target))
                                {
                                    // Only populate missing fields; avoid overwriting values from the selected table
                                    var changedFromBlocks = false;
                                    if (string.IsNullOrWhiteSpace(target.Lat) && !string.IsNullOrWhiteSpace(record.Lat))
                                    {
                                        target.Lat = record.Lat;
                                        changedFromBlocks = true;
                                    }
                                    if (string.IsNullOrWhiteSpace(target.Long) && !string.IsNullOrWhiteSpace(record.Long))
                                    {
                                        target.Long = record.Long;
                                        changedFromBlocks = true;
                                    }
                                    if (string.IsNullOrWhiteSpace(target.Zone) && !string.IsNullOrWhiteSpace(record.Zone))
                                    {
                                        target.Zone = record.Zone;
                                        changedFromBlocks = true;
                                    }
                                    if (changedFromBlocks)
                                        blockEnriched++;
                                }
                            }
                            Logger.Info(ed, $"enriched_from source=blocks keys={blockEnriched}");
                        }
                        catch (System.Exception exScan)
                        {
                            // Best effort: do not abort the command on scan failures; just log the issue.
                            Logger.Warn(ed, $"enrich_latlong_from_blocks_failed err={exScan.Message}");
                        }
                    }

                    var repository = new XingRepository(doc);

                    // Collect extents of all tables so we can ignore blocks embedded in tables
                    var db = doc.Database;
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var tableExtents = new List<Extents3d>();
                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (!btr.IsLayout) continue;

                        foreach (ObjectId id in btr)
                        {
                            var t = tr.GetObject(id, OpenMode.ForRead) as Table;
                            if (t == null) continue;
                            try { tableExtents.Add(t.GeometricExtents); } catch { }
                        }
                    }

                    int totalXing2 = 0, matched = 0, updated = 0, skippedNoKey = 0, skippedNoMatch = 0, matchedByComposite = 0, errors = 0;

                    // Update blocks in model & paper spaces (but skip those intersecting any table extents)
                    foreach (ObjectId btrId in blockTable)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (!btr.IsLayout) continue;

                        foreach (ObjectId entId in btr)
                        {
                            var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                            if (br == null) continue;

                            bool insideTable = false;
                            try
                            {
                                var ext = br.GeometricExtents;
                                foreach (var tExt in tableExtents)
                                {
                                    bool xOver = ext.MinPoint.X <= tExt.MaxPoint.X && ext.MaxPoint.X >= tExt.MinPoint.X;
                                    bool yOver = ext.MinPoint.Y <= tExt.MaxPoint.Y && ext.MaxPoint.Y >= tExt.MinPoint.Y;
                                    if (xOver && yOver) { insideTable = true; break; }
                                }
                            }
                            catch { }
                            if (insideTable) continue;

                            var name = GetEffectiveBlockName(br, tr);
                            if (!string.Equals(name, "XING2", StringComparison.OrdinalIgnoreCase))
                                continue;

                            totalXing2++;
                            try
                            {
                                ProcessBlock(
                                    ed, br, tr, tableType,
                                    byKey, byComposite, byCompositeXKey, repository,
                                    ref matched, ref updated, ref skippedNoKey, ref skippedNoMatch, ref matchedByComposite);
                            }
                            catch (System.Exception ex)
                            {
                                errors++;
                                Logger.Error(ed, $"block handle={br.Handle} err={ex.Message}");
                            }
                        }
                    }

                    tr.Commit();

                    Logger.Info(ed, $"summary xing2_total={totalXing2} matched_by_x={matched} updated={updated} skipped_no_x={skippedNoKey} skipped_no_match={skippedNoMatch} errors={errors}");
                }
                }
                catch (System.Exception ex)
                {
                    Logger.Error(ed, $"match_table_failed err={ex.Message}");
                }
            }
        }

        private static void ProcessBlock(
            Editor ed,
            BlockReference br,
            Transaction tr,
            TableSync.XingTableType tableType,
            IDictionary<string, CrossingRecord> byKey,
            IDictionary<string, CrossingRecord> byComposite,
            IDictionary<string, string> byCompositeXKey,
            XingRepository repository,
            ref int matched,
            ref int updated,
            ref int skippedNoKey,
            ref int skippedNoMatch,
            ref int matchedByComposite)
        {
            if (br == null) return;

            var handle = br.Handle.ToString();
            var keyValue = GetAttributeText(br, tr, CrossingAttributeTags);
            var normalizedKey = TableSync.NormalizeKeyForLookup(keyValue); // same normalizer as tables.

            CrossingRecord record = null;

            if (!string.IsNullOrEmpty(normalizedKey))
            {
                byKey?.TryGetValue(normalizedKey, out record);
            }

            // IMPORTANT:
            // Match Table is intended to update text fields for an existing, correct X#.
            // Fallback/composite matching is intentionally disabled because it can produce
            // unpredictable results in older drawings (duplicates, blanks, etc.).

            if (record == null)
            {
                if (string.IsNullOrEmpty(normalizedKey))
                {
                    skippedNoKey++;
                    Logger.Info(ed, $"block handle={handle} skip reason=no_key");
                }
                else
                {
                    skippedNoMatch++;
                    Logger.Info(ed, $"block handle={handle} skip reason=no_match key={keyValue}");
                }
                return;
            }

            matched++;
            br.UpgradeOpen();

            var changed = false;

            // IMPORTANT: Match Table does NOT change the X# / CROSSING value.
            // The block is matched by its existing X#, and only the other text fields are updated.

            // From table to block
            if (tableType == TableSync.XingTableType.LatLong)
            {
                if (!string.IsNullOrWhiteSpace(record.Description))
                    changed |= SetAttributeIfExists(br, tr, DescriptionAttributeTags, record.Description, null);

                // For LAT/LONG tables, do not write coordinates during MatchTable.
            }
            else
            {
                // MAIN or PAGE table: update textual attributes only
                changed |= SetAttributeIfExists(br, tr, OwnerAttributeTags, record.Owner, null);
                changed |= SetAttributeIfExists(br, tr, DescriptionAttributeTags, record.Description, null);

                if (tableType == TableSync.XingTableType.Main)
                {
                    changed |= SetAttributeIfExists(br, tr, LocationAttributeTags, record.Location, null);
                    changed |= SetAttributeIfExists(br, tr, DwgRefAttributeTags, record.DwgRef, null);
                }

                // DO NOT write LAT/LONG/ZONE values when matching Main or Page tables.
            }

            if (changed)
            {
                updated++;
                br.RecordGraphicsModified(true);
                Logger.Info(ed, $"block handle={handle} updated=true");
            }
            else
            {
                Logger.Debug(ed, $"block handle={handle} updated=false");
            }
        }

        // ---------- Helpers used to build the table indexes ----------
        private static TableSync.XingTableType DetectTableType(Table table, Transaction tr)
        {
            try
            {
                // Reuse your existing IdentifyTable logic.
                var ts = new TableSync(new TableFactory());
                return ts.IdentifyTable(table, tr);
            }
            catch
            {
                return TableSync.XingTableType.Unknown;
            }
        }

        /// <summary>
        /// Reads each data row from the selected table and builds:
        ///  byKey           : "X#" (normalized) -> row record
        ///  byComposite     : composite(owner,desc[,loc,dwg]) -> row record
        ///  byCompositeXKey : composite -> "X#" (normalized)
        /// Tracks duplicate X#s encountered in duplicateKeys.
        /// </summary>
        public static void BuildIndexesFromTable(
            Table table,
            TableSync.XingTableType tableType,
            Editor ed,
            out Dictionary<string, CrossingRecord> byKey,
            out Dictionary<string, CrossingRecord> byComposite,
            out Dictionary<string, string> byCompositeXKey,
            out HashSet<string> duplicateKeys,
            bool logDuplicates = true)
        {
            byKey = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            byComposite = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            byCompositeXKey = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            duplicateKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (table == null) return;

            var rows = table.Rows.Count;
            var cols = table.Columns.Count;

            var dataStartRow = 0;
            if (tableType == TableSync.XingTableType.LatLong)
            {
                var start = TableSync.FindLatLongDataStartRow(table);
                if (start > 0)
                    dataStartRow = start;
            }

            for (int row = dataStartRow; row < rows; row++)
            {
                // Column A: read CROSSING from the block cell (falls back to text).
                var rawKey = TableSync.ResolveCrossingKey(table, row, 0);
                var normalized = TableSync.NormalizeKeyForLookup(rawKey);
                if (string.IsNullOrEmpty(normalized))
                {
                    // Likely header/empty row; ignore.
                    continue;
                }

                // Adjacent cells
                string owner = string.Empty;
                string desc = string.Empty;
                string loc = string.Empty;
                string dwg = string.Empty;
                string lat = string.Empty;
                string lng = string.Empty;
                string zone = string.Empty;

                if (tableType == TableSync.XingTableType.LatLong)
                {
                    desc = (cols > 1) ? ReadCellValue(table, row, 1) : string.Empty;

                    var hasExtendedLayout = cols >= 5;
                    var zoneColumn = hasExtendedLayout ? 2 : -1;
                    var latColumn = hasExtendedLayout ? 3 : (cols > 2 ? 2 : -1);
                    var longColumn = hasExtendedLayout ? 4 : (cols > 3 ? 3 : -1);
                    var dwgColumn = cols >= 6 ? 5 : -1;

                    var zoneLabel = zoneColumn >= 0 ? ReadCellValue(table, row, zoneColumn) : string.Empty;
                    zone = ExtractZoneValue(zoneLabel);
                    lat = latColumn >= 0 ? ReadCellValue(table, row, latColumn) : string.Empty;
                    lng = longColumn >= 0 ? ReadCellValue(table, row, longColumn) : string.Empty;
                    dwg = dwgColumn >= 0 ? ReadCellValue(table, row, dwgColumn) : string.Empty;

                    if (string.IsNullOrWhiteSpace(desc) &&
                        string.IsNullOrWhiteSpace(lat) &&
                        string.IsNullOrWhiteSpace(lng) &&
                        string.IsNullOrWhiteSpace(zone) &&
                        string.IsNullOrWhiteSpace(dwg))
                    {
                        continue;
                    }
                }
                else
                {
                    owner = (cols > 1) ? ReadCellValue(table, row, 1) : string.Empty;
                    desc = (cols > 2) ? ReadCellValue(table, row, 2) : string.Empty;

                    if (tableType == TableSync.XingTableType.Main)
                    {
                        loc = (cols > 3) ? ReadCellValue(table, row, 3) : string.Empty;
                        dwg = (cols > 4) ? ReadCellValue(table, row, 4) : string.Empty;
                    }
                }

                var rec = new CrossingRecord
                {
                    Crossing = (rawKey ?? string.Empty).Trim(),
                    Owner = owner,
                    Description = desc,
                    Location = loc,
                    DwgRef = dwg,
                    Lat = lat,
                    Long = lng,
                    Zone = zone
                };

                if (byKey.ContainsKey(normalized))
                {
                    duplicateKeys.Add(normalized);
                }
                else
                {
                    byKey[normalized] = rec;
                }

                // Composite key (to match blocks even when X is off)
                if (tableType != TableSync.XingTableType.LatLong)
                {
                    var composite = (tableType == TableSync.XingTableType.Main)
                        ? CompositeKeyMain(owner, desc, loc, dwg)
                        : CompositeKeyPage(owner, desc);

                    if (!string.IsNullOrWhiteSpace(composite) && !byComposite.ContainsKey(composite))
                    {
                        byComposite[composite] = rec;
                        byCompositeXKey[composite] = normalized;
                    }
                }
            }

            if (logDuplicates && duplicateKeys.Count > 0)
                Logger.Warn(ed, $"table_duplicate_keys count={duplicateKeys.Count} keys={string.Join(",", duplicateKeys)}");
        }

        // ---------- Local cell/attribute helpers ----------
        private static string GetAttributeText(BlockReference br, Transaction tr, ISet<string> tags)
        {
            if (br == null || tags == null || br.AttributeCollection == null)
                return string.Empty;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                if (!attId.IsValid) continue;
                var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (attRef == null) continue;
                if (tags.Contains(attRef.Tag))
                    return attRef.TextString;
            }
            return string.Empty;
        }

        private static bool SetAttributeIfExists(BlockReference br, Transaction tr, ISet<string> tags, string value, ISet<string> onlyIfMissing)
        {
            bool changed = false;
            if (br.AttributeCollection == null) return false;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (attRef == null) continue;
                if (!tags.Contains(attRef.Tag)) continue;

                if (onlyIfMissing != null && onlyIfMissing.Contains(attRef.Tag) && !string.IsNullOrEmpty(attRef.TextString))
                    continue;

                var desired = value ?? string.Empty;
                if (!string.Equals(attRef.TextString, desired, StringComparison.Ordinal))
                {
                    attRef.TextString = desired;
                    changed = true;
                }
            }
            return changed;
        }

        private static string ExtractZoneValue(string zoneLabel)
        {
            if (string.IsNullOrWhiteSpace(zoneLabel))
                return string.Empty;

            var normalized = TableSync.NormalizeText(zoneLabel) ?? string.Empty;
            normalized = normalized.Trim();
            if (normalized.Length == 0)
                return string.Empty;

            var match = Regex.Match(normalized, "(\\d+)");
            if (match.Success)
                return match.Groups[1].Value.TrimStart('0');

            if (normalized.StartsWith("ZONE", StringComparison.OrdinalIgnoreCase))
            {
                var remainder = normalized.Substring(4).Trim();
                return remainder.Length > 0 ? remainder : string.Empty;
            }

            return normalized;
        }

        private static bool ValuesEqual(string left, string right)
        {
            var leftNorm = (left ?? string.Empty).Trim();
            var rightNorm = (right ?? string.Empty).Trim();
            return string.Equals(leftNorm, rightNorm, StringComparison.OrdinalIgnoreCase);
        }

        private static string GetEffectiveBlockName(BlockReference br, Transaction tr)
        {
            if (br == null) return string.Empty;
            var btrId = br.DynamicBlockTableRecord != ObjectId.Null ? br.DynamicBlockTableRecord : br.BlockTableRecord;
            var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
            return btr.Name;
        }

        // Treat solitary "-" (and em-dash) as empty placeholders
        private static string ReadCellValue(Table table, int row, int column)
        {
            if (table == null || row < 0 || column < 0) return string.Empty;
            if (row >= table.Rows.Count || column >= table.Columns.Count) return string.Empty;

            try
            {
                var cell = table.Cells[row, column];
                var text = cell?.TextString ?? string.Empty;
                return (text ?? string.Empty).Trim();
            }
            catch { return string.Empty; }
        }

    }
}

/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\TableSync.cs
/////////////////////////////////////////////////////////////////////

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
            "XNUM",
            "XNUMBER",
            "NUMBER",
            "INDEX",
            "NO",
            "LABEL"
        };

        private const int MaxHeaderRowsToScan = 5;

        private static readonly Regex MTextFormattingCommandRegex = new Regex("\\\\[^;\\\\{}]*;", RegexOptions.Compiled);
        private static readonly Regex MTextResidualCommandRegex = new Regex("\\\\[^{}]", RegexOptions.Compiled);
        private static readonly Regex MTextSpecialCodeRegex = new Regex("%%[^\\s]+", RegexOptions.Compiled);

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
                                        UpdatePageTable(table, byKey, out matched, out updated);
                                        break;
                                    case XingTableType.LatLong:
                                        var rowIndexMap = BuildLatLongRowIndexMap(table.ObjectId, byKey?.Values);
                                        UpdateLatLongTable(table, byKey, out matched, out updated, rowIndexMap);
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
                                        UpdatePageTable(table, byKey, out matched, out updated);
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

                    int matched = 0, updated = 0;

                    try
                    {
                        if (identifiedType == XingTableType.LatLong)
                        {
                            // Normal header-driven updater (program-generated tables)
                            UpdateLatLongTable(table, byKey, out matched, out updated, rowIndexMap);
                        }
                        else
                        {
                            // Try normal pass anyway
                            UpdateLatLongTable(table, byKey, out matched, out updated, rowIndexMap);

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
                .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();

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
                        TrySetCellBlockScale(t, row, 0, 1.0);   // force block scale = 1.0
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

            var boldStyleId = EnsureBoldTextStyle(db, tr, "XING_BOLD", "Standard");
            var headerColor = Color.FromColorIndex(ColorMethod.ByAci, 254);
            var titleColor = Color.FromColorIndex(ColorMethod.ByAci, 14);

            if (showColumnHeaders)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    var cell = table.Cells[headerRow, c];
                    cell.TextHeight = CellTextHeight;
                    cell.Alignment = CellAlignment.MiddleCenter;
                    cell.TextStyleId = boldStyleId;
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
                titleCell.TextStyleId = boldStyleId;
                ApplyCellTextColor(titleCell, titleColor);
                ApplyTitleBorderStyle(table, titleRow, 0, table.Columns.Count - 1);
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
                    headerCell.TextStyleId = boldStyleId;
                    headerCell.TextString = headerText;
                    ApplyCellTextColor(headerCell, titleColor);
                    ApplyTitleBorderStyle(table, row, 0, table.Columns.Count - 1);
                    continue;
                }

                if (rowInfo.IsColumnHeader)
                {
                    ApplyLatLongColumnHeaderRow(table, row, HeaderRowHeight, CellTextHeight, boldStyleId, headerColor);
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
                SetBorderVisible(bordersObj, "Bottom", true);
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
                                if (string.Equals(text, "LATLONG", StringComparison.OrdinalIgnoreCase)) return XingTableType.LatLong;
                            }
                        }
                    }
                }
            }

            // Header-free heuristics: MAIN=5 cols, PAGE=3 cols, LATLONG=4 cols (+ sanity)
            if (table.Columns.Count == 5) return XingTableType.Main;
            if (table.Columns.Count == 3) return XingTableType.Page;

            if ((table.Columns.Count == 4 || table.Columns.Count >= 6) && table.Rows.Count >= 1)
            {
                var headerColumns = Math.Min(table.Columns.Count, 6);
                if (HasHeaderRow(table, headerColumns, IsLatLongHeader) || LooksLikeLatLongTable(table))
                    return XingTableType.LatLong;
            }

            return XingTableType.Unknown;
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

        private void UpdatePageTable(Table table, IDictionary<string, CrossingRecord> byKey, out int matched, out int updated)
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

                if (record == null) record = FindRecordByPageColumns(table, row, records);

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

                var rowUpdated = false;

                if (columnCount > 0 && ValueDiffers(rawKey, desiredCrossing)) rowUpdated = true;
                if (columnCount > 1 && ValueDiffers(ReadCellText(table, row, 1), desiredOwner)) rowUpdated = true;
                if (columnCount > 2 && ValueDiffers(ReadCellText(table, row, 2), desiredDescription)) rowUpdated = true;

                if (columnCount > 0) SetCellCrossingValue(table, row, 0, desiredCrossing);

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

            RefreshTable(table);
        }

        // TableSync.cs
        private void UpdateLatLongTable(
            Table table,
            IDictionary<string, CrossingRecord> byKey,
            out int matched,
            out int updated,
            IDictionary<int, CrossingRecord> rowIndexMap = null)
        {
            matched = 0;
            updated = 0;
            if (table == null) return;

            var columnCount = table.Columns.Count;
            var records = byKey?.Values ?? Enumerable.Empty<CrossingRecord>();
            var recordList = records.Where(r => r != null).ToList();

            var crossDescMap = BuildCrossingDescriptionMap(recordList);
            var descriptionMap = BuildDescriptionMap(recordList);

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

                // IMPORTANT:
                // - For LAT/LONG tables, the row identity is the cell contents (description + coordinates),
                //   NOT the bubble value. The bubble is only used to read the existing X# for comparison/update.
                var descriptionText = ReadCellTextSafe(table, row, 1);
                var descriptionKey = NormalizeDescriptionKey(descriptionText);

                CrossingRecord record = null;

                // 1) Prefer stable row->record mapping, but only if it still matches this row's description.
                if (rowIndexMap != null && rowIndexMap.TryGetValue(row, out var mapRecord))
                {
                    if (string.IsNullOrEmpty(descriptionKey) ||
                        string.Equals(NormalizeDescriptionKey(mapRecord.Description), descriptionKey, StringComparison.Ordinal))
                    {
                        record = mapRecord;
                    }
                }

                // 2) Primary match = table cell contents (description + lat/long [+ zone/dwg])
                if (record == null)
                {
                    record = FindRecordByLatLongColumns(table, row, recordList);
                }

                // 3) If description is unique, it is safe to match on that alone.
                if (record == null && !string.IsNullOrEmpty(descriptionKey))
                {
                    if (descriptionMap.TryGetValue(descriptionKey, out var list) && list.Count == 1)
                    {
                        record = list[0];
                    }
                }

                // 4) Last resort: use existing X# from the bubble, but ONLY if it matches the row description.
                if (record == null && !string.IsNullOrEmpty(key) && byKey.TryGetValue(key, out var keyed))
                {
                    if (string.IsNullOrEmpty(descriptionKey) ||
                        string.Equals(NormalizeDescriptionKey(keyed.Description), descriptionKey, StringComparison.Ordinal))
                    {
                        record = keyed;
                    }
                }

                if (record == null)
                {
                    // Logger.Debug in this project requires (Editor, message).
                    Logger.Debug(_ed, $"Table {table.ObjectId.Handle}: no match for lat/long row {row} (x='{rawKey}', desc='{descriptionText}')");
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

        private static bool IsCoordinateValue(string text, double min, double max)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            if (!double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var value)) return false;
            return value >= min && value <= max;
        }

        private static string ReadCellText(Cell cell)
        {
            if (cell == null) return string.Empty;
            try { return cell.TextString ?? string.Empty; } catch { return string.Empty; }
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
        // Ensures autoscale is OFF and final scale == 1.0
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
                // Force scale to 1 after placement (tableâ€‘level API if available, else perâ€‘content)
                TrySetCellBlockScale(table, row, col, 1.0);
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

            // 1) Try to set the block attribute by any of our known tags
            foreach (var tag in CrossingAttributeTags)
            {
                if (TrySetBlockAttributeValue(t, row, col, tag, crossingText))
                    return;
            }

            // 2) If the cell currently hosts a block, do NOT overwrite it with text.
            Cell cell = null;
            try { cell = t.Cells[row, col]; } catch { cell = null; }
            if (CellHasBlockContent(cell)) return;

            // 3) Fallback to plain text when no block
            try
            {
                if (cell != null) cell.TextString = crossingText ?? string.Empty;
            }
            catch { }
        }

        private static void SetCellValue(Cell cell, string value)
        {
            if (cell == null) return;

            var desired = value ?? string.Empty;

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
        // Try to set the scale of a block hosted in a table cell to a specific value (e.g., 1.0)
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
            if (candidates.Count == 1) return candidates[0];

            candidates = records.Where(r => string.Equals(Norm(r.Description), desc, StringComparison.Ordinal)).ToList();
            return candidates.Count == 1 ? candidates[0] : null;
        }

        private static CrossingRecord FindRecordByLatLongColumns(Table t, int row, IEnumerable<CrossingRecord> records)
        {
            var desc = ReadNorm(t, row, 1);
            var columns = t?.Columns.Count ?? 0;
            var zone = columns >= 6 ? ReadNorm(t, row, 2) : string.Empty;
            var lat = ReadNorm(t, row, columns >= 6 ? 3 : 2);
            var lng = ReadNorm(t, row, columns >= 6 ? 4 : 3);
            var dwg = columns >= 6 ? ReadNorm(t, row, 5) : string.Empty;

            var candidates = records.Where(r =>
                string.Equals(Norm(r.Description), desc, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Lat), lat, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Long), lng, StringComparison.Ordinal) &&
                (string.IsNullOrEmpty(zone) || string.Equals(Norm(r.Zone), zone, StringComparison.Ordinal)) &&
                (string.IsNullOrEmpty(dwg) || string.Equals(Norm(r.DwgRef), dwg, StringComparison.Ordinal))).ToList();
            if (candidates.Count == 1) return candidates[0];

            candidates = records.Where(r => string.Equals(Norm(r.Description), desc, StringComparison.Ordinal)).ToList();
            return candidates.Count == 1 ? candidates[0] : null;
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

/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\Services\XingRepository.cs
/////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using XingManager.Models;

namespace XingManager.Services
{
    /// <summary>
    /// Provides scan and persistence helpers for xing2 blocks.
    /// </summary>
    public class XingRepository
    {
        public const string BlockName = "xing2";
        public const string LatLongDictionaryKey = "XING2_LATLNG";

        private readonly Document _doc;

        public XingRepository(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        public Document Document => _doc;

        /// <summary>
        /// Scans the drawing for xing2 blocks and collects their attribute values.
        /// Block references inserted into table cells are ignored by checking their
        /// geometric extents against the bounding extents of each table in the drawing.
        /// </summary>
        public ScanResult ScanCrossings()
        {
            var db = _doc.Database;
            var records = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            var contexts = new Dictionary<ObjectId, DuplicateResolver.InstanceContext>();
            var latLongRows = new List<LatLongRowInfo>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Collect extents of all tables in the drawing.
                var tableExtents = new List<TableExtentInfo>();
                var crossingTableRowsByTableId = new Dictionary<ObjectId, Dictionary<string, CrossingTableRowData>>();
                var tableBubbleBlocks = new Dictionary<ObjectId, ObjectId>();
                var tableOwnerBtrByTableId = new Dictionary<ObjectId, ObjectId>();
                var tableHandleByTableId = new Dictionary<ObjectId, string>();
                var btab = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in btab)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    if (!btr.IsLayout) continue;

                    foreach (ObjectId entId in btr)
                    {
                        var tbl = tr.GetObject(entId, OpenMode.ForRead) as Table;
                        if (tbl == null) continue;

                        // Track which layout/paperspace this table belongs to (for duplicate resolver labels)
                        tableOwnerBtrByTableId[tbl.ObjectId] = btrId;
                        try { tableHandleByTableId[tbl.ObjectId] = tbl.Handle.ToString(); } catch { /* ignore */ }


                        try
                        {
                            var ext = tbl.GeometricExtents;
                            tableExtents.Add(new TableExtentInfo { TableId = tbl.ObjectId, Extents = ext });
                        }
                        catch
                        {
                            // ignore tables without extents
                        }

                        try
                        {
                            CollectLatLongRows(tbl, latLongRows);

                            var rowMap = CollectCrossingTableRowMap(tbl);
                            if (rowMap.Count > 0)
                            {
                                crossingTableRowsByTableId[tbl.ObjectId] = rowMap;
                            }
                        }
                        catch
                        {
                            // ignore lat/long parsing errors
                        }
                    }
                }

                // Build a map of layout BlockTableRecordId -> layout name.
                var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                var layoutNames = new Dictionary<ObjectId, string>();
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    layoutNames[layout.BlockTableRecordId] = layout.LayoutName;
                }

                var blockRefClass = RXClass.GetClass(typeof(BlockReference));

                foreach (ObjectId btrId in btab)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    if (!btr.IsLayout)
                        continue;

                    var spaceName = layoutNames.ContainsKey(btrId) ? layoutNames[btrId] : btr.Name;

                    foreach (ObjectId entId in btr)
                    {
                        if (!entId.ObjectClass.IsDerivedFrom(blockRefClass))
                            continue;

                        var br = (BlockReference)tr.GetObject(entId, OpenMode.ForRead);
                        if (br == null)
                            continue;

                        var blockEffectiveName = GetBlockName(br, tr);
                        if (!string.Equals(blockEffectiveName, BlockName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        // Determine if this block is inside a table or on a paper-space layout.
                        bool isInTable = false;
                        ObjectId tableId = ObjectId.Null;
                        try
                        {
                            var blkExt = br.GeometricExtents;
                            foreach (var te in tableExtents)
                            {
                                var tblExt = te.Extents;
                                bool xOverlaps = blkExt.MinPoint.X <= tblExt.MaxPoint.X && blkExt.MaxPoint.X >= tblExt.MinPoint.X;
                                bool yOverlaps = blkExt.MinPoint.Y <= tblExt.MaxPoint.Y && blkExt.MaxPoint.Y >= tblExt.MinPoint.Y;
                                if (xOverlaps && yOverlaps)
                                {
                                    isInTable = true;
                                    tableId = te.TableId;
                                    break;
                                }
                            }
                        }
                        catch
                        {
                            // ignore extents errors
                        }

                        // Capture attributes
                        var attributes = ReadAttributes(br, tr);
                        var crossing = SanitizeAttributeValue(GetValue(attributes, "CROSSING"));
                        if (string.IsNullOrEmpty(crossing))
                            continue;

                        var owner = SanitizeAttributeValue(GetValue(attributes, "OWNER"));
                        var description = SanitizeAttributeValue(GetValue(attributes, "DESCRIPTION"));
                        var location = SanitizeAttributeValue(GetValue(attributes, "LOCATION"));
                        var dwgRef = SanitizeAttributeValue(GetValue(attributes, "DWG_REF"));
                        string lat;
                        string lng;
                        string zone;
                        TryGetLatLong(br, tr, out lat, out lng, out zone);

                        var crossingKey = crossing.Trim().ToUpperInvariant();
                        CrossingRecord record;
                        if (!records.TryGetValue(crossingKey, out record))
                        {
                            record = new CrossingRecord
                            {
                                Crossing = crossing,
                                Owner = owner,
                                Description = description,
                                Location = location,
                                DwgRef = dwgRef,
                                Lat = lat,
                                Long = lng,
                                Zone = zone,
                                CanonicalInstance = ObjectId.Null
                            };
                            records.Add(crossingKey, record);
                        }

                        record.AllInstances.Add(entId);

                        // Create instance context; mark as ignored if in table or not in Model space
                        bool ignore = isInTable || !string.Equals(spaceName, "Model", StringComparison.OrdinalIgnoreCase);
                        contexts[entId] = new DuplicateResolver.InstanceContext
                        {
                            ObjectId = entId,
                            Crossing = crossing,
                            SpaceName = spaceName,
                            Owner = owner,
                            Description = description,
                            Location = location,
                            DwgRef = dwgRef,
                            Lat = lat,
                            Long = lng,
                            Zone = zone,
                            IgnoreForDuplicates = ignore,
                            IsTableInstance = isInTable
                        };

                        if (isInTable && tableId != ObjectId.Null && crossingTableRowsByTableId.ContainsKey(tableId))
                        {
                            tableBubbleBlocks[entId] = tableId;
                        }

                        // Prefer a model-space instance as canonical; otherwise choose first.
                        if (record.CanonicalInstance.IsNull &&
                            string.Equals(spaceName, "Model", StringComparison.OrdinalIgnoreCase))
                        {
                            record.CanonicalInstance = entId;
                            record.Crossing = crossing;
                            record.Owner = owner;
                            record.Description = description;
                            record.Location = location;
                            record.DwgRef = dwgRef;
                            record.Lat = lat;
                            record.Long = lng;
                            record.Zone = zone;
                        }
                        else if (record.CanonicalInstance.IsNull)
                        {
                            record.CanonicalInstance = entId;
                            record.Owner = owner;
                            record.Description = description;
                            record.Location = location;
                            record.DwgRef = dwgRef;
                            record.Lat = lat;
                            record.Long = lng;
                            record.Zone = zone;
                        }
                    }
                }

	                // Apply table-row values into per-instance contexts so the duplicate resolver can
	                // surface mismatches between drawing blocks and table rows.
	                ApplyCrossingTableOverrides(tableBubbleBlocks, crossingTableRowsByTableId, records, contexts);

	                // Attach crossing table row sources to the corresponding records so mismatches with
	                // block/UI values can trigger the crossing duplicate resolver.
	                ApplyCrossingTableRowsToRecords(crossingTableRowsByTableId, tableOwnerBtrByTableId, tableHandleByTableId, layoutNames, records);

	                tr.Commit();
            }

	            ApplyLatLongRows(latLongRows, records);

            var ordered = records.Values
                .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();

            return new ScanResult
            {
                Records = ordered,
                InstanceContexts = contexts
            };
        }

        public void ApplyChanges(IList<CrossingRecord> records, TableSync tableSync)
        {
            if (records == null) throw new ArgumentNullException(nameof(records));
            if (tableSync == null) throw new ArgumentNullException(nameof(tableSync));

            var db = _doc.Database;

            using (_doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var record in records)
                {
                    foreach (var instanceId in record.AllInstances.Distinct())
                    {
                        if (!instanceId.IsValid)
                            continue;

                        var br = tr.GetObject(instanceId, OpenMode.ForWrite) as BlockReference;
                        if (br == null)
                            continue;

                        // Write block attributes from the in-memory record
                        WriteAttribute(tr, br, "CROSSING", record.Crossing);
                        WriteAttribute(tr, br, "OWNER", record.Owner);
                        WriteAttribute(tr, br, "DESCRIPTION", record.Description);
                        WriteAttribute(tr, br, "LOCATION", record.Location);
                        WriteAttribute(tr, br, "DWG_REF", record.DwgRef);
                        SetLatLong(br, tr, record.Lat, record.Long, record.Zone);
                    }
                }

                tr.Commit();
            }

            // Automatic table updates were previously removed (see historical comment: "IMPORTANT: removed automatic table updates here.");
            // Restore that behavior so every attribute write immediately syncs the LAT/LONG and other crossing tables without extra commands.
            tableSync.UpdateLatLongSourceTables(_doc, records);
            tableSync.UpdateAllTables(_doc, records);
        }

        public ObjectId InsertCrossing(CrossingRecord record, Point3d position)
        {
            if (record == null)
                throw new ArgumentNullException(nameof(record));

            var db = _doc.Database;

            using (_doc.LockDocument())
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    if (!blockTable.Has(BlockName))
                        throw new InvalidOperationException("Block 'xing2' definition not found in this drawingÃ¢â‚¬â€insert one or use INSERT to bring it in, then retry.");

                    var blockId = blockTable[BlockName];
                    var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    var blockDef = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);

                    var br = new BlockReference(position, blockId);
                    modelSpace.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    if (blockDef.HasAttributeDefinitions)
                    {
                        foreach (ObjectId attDefId in blockDef)
                        {
                            var attDef = tr.GetObject(attDefId, OpenMode.ForRead) as AttributeDefinition;
                            if (attDef == null || attDef.Constant)
                                continue;

                            var attRef = new AttributeReference();
                            attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                            attRef.TextString = attDef.TextString;
                            br.AttributeCollection.AppendAttribute(attRef);
                            tr.AddNewlyCreatedDBObject(attRef, true);
                        }
                    }

                    WriteAttribute(tr, br, "CROSSING", record.Crossing);
                    WriteAttribute(tr, br, "OWNER", record.Owner);
                    WriteAttribute(tr, br, "DESCRIPTION", record.Description);
                    WriteAttribute(tr, br, "LOCATION", record.Location);
                    WriteAttribute(tr, br, "DWG_REF", record.DwgRef);
                    SetLatLong(br, tr, record.Lat, record.Long, record.Zone);

                    tr.Commit();
                    return br.ObjectId;
                }
            }
        }

        public void DeleteInstances(IEnumerable<ObjectId> instanceIds)
        {
            if (instanceIds == null) return;

            using (_doc.LockDocument())
            using (var tr = _doc.Database.TransactionManager.StartTransaction())
            {
                foreach (var id in instanceIds.Distinct())
                {
                    if (!id.IsValid) continue;

                    var dbObject = tr.GetObject(id, OpenMode.ForWrite, false, true);
                    if (dbObject != null && !dbObject.IsErased)
                        dbObject.Erase(true);
                }

                tr.Commit();
            }
        }

        public void DeleteBlocksByCrossing(string crossing)
        {
            if (string.IsNullOrWhiteSpace(crossing))
                return;

            var targetKey = crossing.Trim().ToUpperInvariant();

            using (_doc.LockDocument())
            using (var tr = _doc.Database.TransactionManager.StartTransaction())
            {
                var db = _doc.Database;
                var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var blockRefClass = RXClass.GetClass(typeof(BlockReference));
                var idsToDelete = new List<ObjectId>();

                foreach (ObjectId btrId in blockTable)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    if (!btr.IsLayout)
                        continue;

                    foreach (ObjectId entId in btr)
                    {
                        if (!entId.ObjectClass.IsDerivedFrom(blockRefClass))
                            continue;

                        var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                        if (br == null || br.IsErased)
                            continue;

                        var blockName = GetBlockName(br, tr);
                        if (!string.Equals(blockName, BlockName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        var attributes = ReadAttributes(br, tr);
                        var attValue = GetValue(attributes, "CROSSING");
                        var attKey = (attValue ?? string.Empty).Trim().ToUpperInvariant();
                        if (string.IsNullOrEmpty(attKey))
                            continue;

                        if (string.Equals(attKey, targetKey, StringComparison.Ordinal))
                            idsToDelete.Add(entId);
                    }
                }

                foreach (var id in idsToDelete.Distinct())
                {
                    var br = tr.GetObject(id, OpenMode.ForWrite, false, true) as BlockReference;
                    if (br != null && !br.IsErased)
                        br.Erase(true);
                }

                tr.Commit();
            }
        }

        public bool TryGetLatLong(BlockReference br, Transaction tr, out string lat, out string lng, out string zone)
        {
            lat = string.Empty;
            lng = string.Empty;
            zone = string.Empty;
            if (br == null) return false;

            if (br.ExtensionDictionary.IsNull)
                return false;

            var dict = tr.GetObject(br.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
            if (dict == null || !dict.Contains(LatLongDictionaryKey))
                return false;

            var xrec = tr.GetObject(dict.GetAt(LatLongDictionaryKey), OpenMode.ForRead) as Xrecord;
            if (xrec == null || xrec.Data == null)
                return false;

            var values = xrec.Data.AsArray();
            if (values.Length >= 1)
                lat = Convert.ToString(values[0].Value, CultureInfo.InvariantCulture);

            if (values.Length >= 2)
                lng = Convert.ToString(values[1].Value, CultureInfo.InvariantCulture);

            if (values.Length >= 3)
                zone = Convert.ToString(values[2].Value, CultureInfo.InvariantCulture);

            return !string.IsNullOrEmpty(lat) || !string.IsNullOrEmpty(lng) || !string.IsNullOrEmpty(zone);
        }

        public void SetLatLong(BlockReference br, Transaction tr, string lat, string lng, string zone)
        {
            if (br == null) throw new ArgumentNullException(nameof(br));
            if (tr == null) throw new ArgumentNullException(nameof(tr));

            if (br.ExtensionDictionary.IsNull)
                br.CreateExtensionDictionary();

            var dict = (DBDictionary)tr.GetObject(br.ExtensionDictionary, OpenMode.ForWrite);
            Xrecord xrec;
            if (dict.Contains(LatLongDictionaryKey))
                xrec = (Xrecord)tr.GetObject(dict.GetAt(LatLongDictionaryKey), OpenMode.ForWrite);
            else
            {
                xrec = new Xrecord();
                dict.SetAt(LatLongDictionaryKey, xrec);
                tr.AddNewlyCreatedDBObject(xrec, true);
            }

            xrec.Data = new ResultBuffer(
                new TypedValue((int)DxfCode.Text, lat ?? string.Empty),
                new TypedValue((int)DxfCode.Text, lng ?? string.Empty),
                new TypedValue((int)DxfCode.Text, zone ?? string.Empty));
        }

        internal static void CollectLatLongRows(Table table, IList<LatLongRowInfo> rows)
        {
            if (table == null || rows == null)
                return;

            if (!IsLatLongTable(table))
                return;

            var startRow = TableSync.FindLatLongDataStartRow(table);
            if (startRow <= 0)
            {
                startRow = 0;
            }

            var columnCount = table.Columns.Count;
            var hasZoneColumn = columnCount >= 6;
            var hasDwgColumn = columnCount >= 6;

            var zoneColumn = hasZoneColumn ? 2 : -1;
            var latColumn = hasZoneColumn ? 3 : 2;
            var longColumn = hasZoneColumn ? 4 : 3;
            var dwgColumn = hasDwgColumn ? 5 : -1;

            var tableHandle = table.ObjectId.IsNull ? string.Empty : table.ObjectId.Handle.ToString();

            for (var row = startRow; row < table.Rows.Count; row++)
            {
                var crossing = TableSync.ResolveCrossingKey(table, row, 0);
                var description = TableSync.ReadCellTextSafe(table, row, 1);
                var zoneLabel = zoneColumn >= 0 ? TableSync.ReadCellTextSafe(table, row, zoneColumn) : string.Empty;
                var latitude = TableSync.ReadCellTextSafe(table, row, latColumn);
                var longitude = TableSync.ReadCellTextSafe(table, row, longColumn);
                var dwgRef = dwgColumn >= 0 ? TableSync.ReadCellTextSafe(table, row, dwgColumn) : string.Empty;

                if (string.IsNullOrWhiteSpace(crossing) &&
                    string.IsNullOrWhiteSpace(description) &&
                    string.IsNullOrWhiteSpace(latitude) &&
                    string.IsNullOrWhiteSpace(longitude) &&
                    string.IsNullOrWhiteSpace(zoneLabel) &&
                    string.IsNullOrWhiteSpace(dwgRef))
                {
                    continue;
                }

                rows.Add(new LatLongRowInfo
                {
                    TableId = table.ObjectId,
                    RowIndex = row,
                    SourceLabel = string.IsNullOrWhiteSpace(tableHandle)
                        ? string.Format(CultureInfo.InvariantCulture, "LAT/LONG table row {0}", row + 1)
                        : string.Format(CultureInfo.InvariantCulture, "LAT/LONG table {0} row {1}", tableHandle, row + 1),
                    Crossing = crossing,
                    Description = description,
                    Latitude = latitude,
                    Longitude = longitude,
                    Zone = ExtractZoneValue(zoneLabel),
                    DwgRef = dwgRef
                });
            }
        }

        internal static void ApplyLatLongRows(IEnumerable<LatLongRowInfo> rows, IDictionary<string, CrossingRecord> records)
        {
            if (rows == null || records == null)
                return;

            foreach (var record in records.Values)
            {
                record?.LatLongSources?.Clear();
            }

            foreach (var row in rows)
            {
                if (row == null)
                    continue;

                var record = FindRecordForLatLong(row, records);
                if (record == null)
                    continue;

                record.LatLongSources.Add(new CrossingRecord.LatLongSource
                {
                    SourceLabel = row.SourceLabel,
                    Description = row.Description,
                    Lat = row.Latitude,
                    Long = row.Longitude,
                    Zone = row.Zone,
                    DwgRef = row.DwgRef,
                    TableId = row.TableId,
                    RowIndex = row.RowIndex
                });

                var latitude = row.Latitude?.Trim();
                var longitude = row.Longitude?.Trim();
                var zone = row.Zone?.Trim();

                if (!string.IsNullOrWhiteSpace(latitude))
                {
                    record.Lat = latitude;
                }

                if (!string.IsNullOrWhiteSpace(longitude))
                {
                    record.Long = longitude;
                }

                if (!string.IsNullOrWhiteSpace(zone))
                {
                    record.Zone = zone;
                }

                var dwgRef = row.DwgRef?.Trim();
                if (!string.IsNullOrWhiteSpace(dwgRef) && string.IsNullOrWhiteSpace(record.DwgRef))
                {
                    record.DwgRef = dwgRef;
                }
            }

            foreach (var record in records.Values)
            {
                if (record == null || record.LatLongSources == null || record.LatLongSources.Count == 0)
                    continue;

                var normalizedLat = NormalizeForComparison(record.Lat);
                var normalizedLong = NormalizeForComparison(record.Long);
                var normalizedZone = NormalizeForComparison(record.Zone);

                var matching = record.LatLongSources.FirstOrDefault(s =>
                    string.Equals(NormalizeForComparison(s.Lat), normalizedLat, StringComparison.Ordinal) &&
                    string.Equals(NormalizeForComparison(s.Long), normalizedLong, StringComparison.Ordinal) &&
                    (string.IsNullOrEmpty(NormalizeForComparison(s.Zone)) ||
                     string.Equals(NormalizeForComparison(s.Zone), normalizedZone, StringComparison.Ordinal)));

                if (matching != null)
                    continue;

                var candidates = record.LatLongSources
                    .Where(s => !string.IsNullOrWhiteSpace(s.Lat) ||
                                !string.IsNullOrWhiteSpace(s.Long) ||
                                !string.IsNullOrWhiteSpace(s.Zone))
                    .ToList();

                if (candidates.Count == 0)
                    continue;

                var signatures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var candidate in candidates)
                {
                    var signature = string.Join("|",
                        NormalizeForComparison(candidate.Lat),
                        NormalizeForComparison(candidate.Long),
                        NormalizeForComparison(candidate.Zone));
                    signatures.Add(signature);
                }

                if (signatures.Count == 1)
                {
                    var source = candidates[0];
                    if (!string.IsNullOrWhiteSpace(source.Lat))
                        record.Lat = source.Lat.Trim();
                    if (!string.IsNullOrWhiteSpace(source.Long))
                        record.Long = source.Long.Trim();
                    if (!string.IsNullOrWhiteSpace(source.Zone))
                        record.Zone = source.Zone.Trim();
                }
            }
        }

        private static CrossingRecord FindRecordForLatLong(LatLongRowInfo row, IDictionary<string, CrossingRecord> records)
        {
            if (row == null || records == null)
                return null;

            var rawCrossing = row.Crossing ?? string.Empty;
            var normalizedCrossing = rawCrossing.Trim().ToUpperInvariant();
            var normalizedKey = TableSync.NormalizeKeyForLookup(rawCrossing);

            if (!string.IsNullOrWhiteSpace(normalizedCrossing) &&
                records.TryGetValue(normalizedCrossing, out var direct))
            {
                return direct;
            }

            if (!string.IsNullOrWhiteSpace(normalizedKey) &&
                records.TryGetValue(normalizedKey, out var normalizedMatch))
            {
                return normalizedMatch;
            }

            var byComparison = records.Values
                .FirstOrDefault(r => CrossingRecord.CompareCrossingKeys(r?.Crossing, rawCrossing) == 0);
            if (byComparison != null)
            {
                return byComparison;
            }

            var normalizedDescription = NormalizeForComparison(row.Description);
            if (string.IsNullOrEmpty(normalizedDescription))
            {
                return null;
            }

            var candidates = records.Values
                .Where(r => string.Equals(NormalizeForComparison(r?.Description), normalizedDescription, StringComparison.Ordinal))
                .ToList();

            if (candidates.Count == 1)
            {
                return candidates[0];
            }

            if (candidates.Count > 1)
            {
                var latNorm = NormalizeForComparison(row.Latitude);
                var longNorm = NormalizeForComparison(row.Longitude);
                var zoneNorm = NormalizeForComparison(row.Zone);
                var dwgNorm = NormalizeForComparison(row.DwgRef);

                var matches = candidates
                    .Where(r => string.Equals(NormalizeForComparison(r?.Lat), latNorm, StringComparison.Ordinal) &&
                                string.Equals(NormalizeForComparison(r?.Long), longNorm, StringComparison.Ordinal) &&
                                (string.IsNullOrEmpty(zoneNorm) || string.Equals(NormalizeForComparison(r?.Zone), zoneNorm, StringComparison.Ordinal)) &&
                                (string.IsNullOrEmpty(dwgNorm) || string.Equals(NormalizeForComparison(r?.DwgRef), dwgNorm, StringComparison.Ordinal)))
                    .ToList();

                if (matches.Count == 1)
                {
                    return matches[0];
                }
            }

            return null;
        }

        internal static bool IsLatLongTable(Table table)
        {
            if (table == null)
                return false;

            if ((table.Columns.Count != 4 && table.Columns.Count != 6) || table.Rows.Count <= 0)
                return false;

            if (TableSync.FindLatLongDataStartRow(table) > 0)
                return true;

            var normalizedHeaders = new List<string>(4);
            var maxColumns = Math.Min(table.Columns.Count, 6);
            for (var column = 0; column < maxColumns; column++)
            {
                normalizedHeaders.Add(NormalizeForComparison(TableSync.ReadCellTextSafe(table, 0, column)));
            }

            var extendedHeaders = new[] { "ID", "DESCRIPTION", "ZONE", "LATITUDE", "LONGITUDE", "DWG_REF" };
            if (normalizedHeaders.Count == extendedHeaders.Length && normalizedHeaders.SequenceEqual(extendedHeaders))
                return true;

            var updatedHeaders = new[] { "ID", "DESCRIPTION", "LATITUDE", "LONGITUDE" };
            if (normalizedHeaders.Count == updatedHeaders.Length && normalizedHeaders.SequenceEqual(updatedHeaders))
                return true;

            var legacyHeaders = new[] { "XING", "DESCRIPTION", "LAT", "LONG" };
            if (normalizedHeaders.Count == legacyHeaders.Length && normalizedHeaders.SequenceEqual(legacyHeaders))
                return true;

            return LooksLikeLatLongDataRows(table);
        }

        private static string NormalizeForComparison(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            var normalized = TableSync.NormalizeText(value) ?? string.Empty;
            normalized = normalized.Trim();

            while (normalized.Contains("  "))
            {
                normalized = normalized.Replace("  ", " ");
            }

            return normalized.ToUpperInvariant();
        }

        private static bool LooksLikeLatLongDataRows(Table table)
        {
            if (table == null)
                return false;

            var columns = table.Columns.Count;
            if (columns < 4)
                return false;

            var latColumn = columns >= 6 ? 3 : 2;
            var longColumn = columns >= 6 ? 4 : 3;

            var rowCount = table.Rows.Count;
            if (rowCount <= 0)
                return false;

            var rowsToScan = Math.Min(rowCount, 12);
            var candidateRows = 0;

            for (var row = 0; row < rowsToScan; row++)
            {
                var latText = TableSync.ReadCellTextSafe(table, row, latColumn);
                var longText = TableSync.ReadCellTextSafe(table, row, longColumn);

                if (string.IsNullOrWhiteSpace(latText) && string.IsNullOrWhiteSpace(longText))
                    continue;

                var latIsCoordinate = LooksLikeCoordinate(latText, -90.0, 90.0);
                var longIsCoordinate = LooksLikeCoordinate(longText, -180.0, 180.0);

                if (!latIsCoordinate || !longIsCoordinate)
                {
                    if (LooksLikeHeader(latText) && LooksLikeHeader(longText))
                        continue;

                    return false;
                }

                var crossing = TableSync.ResolveCrossingKey(table, row, 0);
                var description = TableSync.ReadCellTextSafe(table, row, 1);

                if (!LooksLikeCrossingValue(crossing) && string.IsNullOrWhiteSpace(description))
                    return false;

                candidateRows++;
            }

            return candidateRows > 0;
        }

        private static bool LooksLikeCoordinate(string text, double min, double max)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var normalized = TableSync.NormalizeText(text);
            if (string.IsNullOrWhiteSpace(normalized))
                return false;

            if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            {
                if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
                    return false;
            }

            return value >= min && value <= max;
        }

        private static bool LooksLikeHeader(string text)
        {
            var normalized = NormalizeForComparison(text);
            if (string.IsNullOrWhiteSpace(normalized))
                return false;

            return normalized.StartsWith("LAT", StringComparison.Ordinal) ||
                   normalized.StartsWith("LONG", StringComparison.Ordinal) ||
                   normalized.Contains("ZONE");
        }

        private static bool LooksLikeCrossingValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            var normalized = NormalizeForComparison(value);
            if (normalized.StartsWith("X", StringComparison.Ordinal))
                return true;

            var token = CrossingRecord.ParseCrossingNumber(value);
            return token.Number > 0;
        }


        private struct TableExtentInfo
        {
            public ObjectId TableId;
            public Extents3d Extents;
        }

        private sealed class CrossingTableRowData
{
    public string Owner = "";
    public string Description = "";
    public string Location = "";
    public string DwgRef = "";
    public bool HasOwner;
    public bool HasLocation;
    public bool HasDwgRef;
    public int RowIndex;
}

        private static Dictionary<string, CrossingTableRowData> CollectCrossingTableRowMap(Table table)
{
    var map = new Dictionary<string, CrossingTableRowData>(StringComparer.OrdinalIgnoreCase);

    if (table == null)
        return map;

    // Exclude LAT/LONG tables; those are handled by the lat/long resolver.
    if (IsLatLongTable(table))
        return map;

    int rowCount;
    int colCount;
    try
    {
        rowCount = table.Rows.Count;
        colCount = table.Columns.Count;
    }
    catch
    {
        return map;
    }

    if (rowCount <= 0 || colCount < 2)
        return map;

    // Supported crossing table shapes:
    //  - Main crossing table: 5 columns (XING / OWNER / DESCRIPTION / LOCATION / DWG_REF)
    //  - Crossing page table: 3 columns (XING / OWNER / DESCRIPTION) -> LOCATION & DWG_REF are ignored when comparing
    //  - Legacy (optional): 4 columns (XING / DESCRIPTION / LOCATION / DWG_REF) -> OWNER ignored when comparing
    bool hasOwner = false;
    bool hasLocation = false;
    bool hasDwgRef = false;
    int ownerCol = -1;
    int descCol = -1;
    int locCol = -1;
    int dwgCol = -1;

    if (colCount >= 5)
    {
        hasOwner = true; ownerCol = 1;
        descCol = 2;
        hasLocation = true; locCol = 3;
        hasDwgRef = true; dwgCol = 4;
    }
    else if (colCount == 3)
    {
        hasOwner = true; ownerCol = 1;
        descCol = 2;
        hasLocation = false;
        hasDwgRef = false;
    }
    else if (colCount == 4)
    {
        // Legacy: no OWNER column
        hasOwner = false;
        descCol = 1;
        hasLocation = true; locCol = 2;
        hasDwgRef = true; dwgCol = 3;
    }
    else
    {
        return map;
    }

    for (int r = 0; r < rowCount; r++)
    {
        string crossing = NormalizeCrossingKey(TableSync.ResolveCrossingKey(table, r, 0));
        if (!LooksLikeCrossingValue(crossing))
            continue;

        var rowData = new CrossingTableRowData
        {
            RowIndex = r,
            HasOwner = hasOwner,
            HasLocation = hasLocation,
            HasDwgRef = hasDwgRef
        };

        if (hasOwner && ownerCol >= 0 && ownerCol < colCount)
            rowData.Owner = TableSync.ReadCellTextSafe(table, r, ownerCol) ?? string.Empty;

        if (descCol >= 0 && descCol < colCount)
            rowData.Description = TableSync.ReadCellTextSafe(table, r, descCol) ?? string.Empty;

        if (hasLocation && locCol >= 0 && locCol < colCount)
            rowData.Location = TableSync.ReadCellTextSafe(table, r, locCol) ?? string.Empty;

        if (hasDwgRef && dwgCol >= 0 && dwgCol < colCount)
            rowData.DwgRef = TableSync.ReadCellTextSafe(table, r, dwgCol) ?? string.Empty;

        map[crossing] = rowData;
    }

    return map;
}

        private static void ApplyCrossingTableOverrides(
            Dictionary<ObjectId, ObjectId> tableBubbleBlocks,
            Dictionary<ObjectId, Dictionary<string, CrossingTableRowData>> crossingTableRowsByTableId,
            Dictionary<string, CrossingRecord> records,
            Dictionary<ObjectId, DuplicateResolver.InstanceContext> contexts)
        {
            if (tableBubbleBlocks == null || crossingTableRowsByTableId == null || records == null || contexts == null)
                return;

            foreach (var kvp in tableBubbleBlocks)
            {
                ObjectId blockId = kvp.Key;
                ObjectId tableId = kvp.Value;

                if (!contexts.TryGetValue(blockId, out var ctx) || ctx == null)
                    continue;

                string crossing = NormalizeCrossingKey(ctx.Crossing);
                if (string.IsNullOrWhiteSpace(crossing))
                    continue;

                if (!records.TryGetValue(crossing, out var rec) || rec == null)
                    continue;

                if (!crossingTableRowsByTableId.TryGetValue(tableId, out var rowMap) || rowMap == null)
                    continue;

                if (!rowMap.TryGetValue(crossing, out var rowData) || rowData == null)
                    continue;

                // For table instances, compare using the cell text (B-D), not the bubble's own attributes.
                ctx.Description = rowData.Description ?? string.Empty;
                ctx.Location = rowData.HasLocation ? (rowData.Location ?? string.Empty) : (rec.Location ?? string.Empty);
                ctx.DwgRef = rowData.HasDwgRef ? (rowData.DwgRef ?? string.Empty) : (rec.DwgRef ?? string.Empty);

                // Keep non-table fields aligned to the record so we don't create false duplicates.
                ctx.Owner = rowData.HasOwner ? (rowData.Owner ?? string.Empty) : (rec.Owner ?? string.Empty);
                ctx.Zone = rec.Zone ?? string.Empty;
                ctx.Lat = rec.Lat ?? string.Empty;
                ctx.Long = rec.Long ?? string.Empty;

                ctx.IgnoreForDuplicates = false;
            }
        }

private static void ApplyCrossingTableRowsToRecords(
    IDictionary<ObjectId, Dictionary<string, CrossingTableRowData>> crossingTableRowsByTableId,
    IDictionary<ObjectId, ObjectId> tableOwnerBtrByTableId,
    IDictionary<ObjectId, string> tableHandleByTableId,
    IDictionary<ObjectId, string> layoutNames,
    IDictionary<string, CrossingRecord> records)
{
    if (records == null)
        return;

    // Reset any previous scan's table sources
    foreach (var rec in records.Values)
    {
        rec?.CrossingTableSources?.Clear();
    }

    if (crossingTableRowsByTableId == null || crossingTableRowsByTableId.Count == 0)
        return;

    foreach (var tableEntry in crossingTableRowsByTableId)
    {
        var tableId = tableEntry.Key;
        var rowMap = tableEntry.Value;
        if (rowMap == null || rowMap.Count == 0)
            continue;

        // Build a helpful label for the resolver (which layout + table handle)
        string layoutLabel = "Unknown";
        if (tableOwnerBtrByTableId != null &&
            tableOwnerBtrByTableId.TryGetValue(tableId, out var ownerBtrId) &&
            layoutNames != null &&
            layoutNames.TryGetValue(ownerBtrId, out var ln) &&
            !string.IsNullOrWhiteSpace(ln))
        {
            layoutLabel = ln;
        }

        string handleLabel = "";
        if (tableHandleByTableId != null &&
            tableHandleByTableId.TryGetValue(tableId, out var handle) &&
            !string.IsNullOrWhiteSpace(handle))
        {
            handleLabel = handle;
        }

        string sourceLabel = string.IsNullOrWhiteSpace(handleLabel)
            ? $"TABLE ({layoutLabel})"
            : $"TABLE ({layoutLabel}) #{handleLabel}";

        foreach (var rowEntry in rowMap)
        {
            var crossingKey = NormalizeCrossingKey(rowEntry.Key);
            if (string.IsNullOrWhiteSpace(crossingKey))
                continue;

            if (!records.TryGetValue(crossingKey, out var rec) || rec == null)
                continue;

            var row = rowEntry.Value;
            if (row == null)
                continue;

            rec.CrossingTableSources.Add(new CrossingRecord.CrossingTableSource
            {
                SourceLabel = sourceLabel,
                Owner = row.Owner ?? string.Empty,
                Description = row.Description ?? string.Empty,
                Location = row.Location ?? string.Empty,
                DwgRef = row.DwgRef ?? string.Empty,
                HasOwner = row.HasOwner,
                HasLocation = row.HasLocation,
                HasDwgRef = row.HasDwgRef,
                TableId = tableId,
                RowIndex = row.RowIndex
            });
        }
    }
}

        private static string NormalizeCrossingKey(string raw)
        {
            return string.IsNullOrWhiteSpace(raw) ? string.Empty : raw.Trim().ToUpperInvariant();
        }

        internal sealed class LatLongRowInfo
        {
            public ObjectId TableId { get; set; }
            public int RowIndex { get; set; }
            public string SourceLabel { get; set; }
            public string Crossing { get; set; }
            public string Description { get; set; }
            public string Latitude { get; set; }
            public string Longitude { get; set; }
            public string Zone { get; set; }
            public string DwgRef { get; set; }
        }

        private static string ExtractZoneValue(string zoneLabel)
        {
            if (string.IsNullOrWhiteSpace(zoneLabel))
            {
                return string.Empty;
            }

            var trimmed = TableSync.NormalizeText(zoneLabel) ?? string.Empty;
            trimmed = trimmed.Trim();
            if (trimmed.Length == 0)
            {
                return string.Empty;
            }

            var match = Regex.Match(trimmed, @"(\d+)");
            if (match.Success)
            {
                return match.Groups[1].Value.TrimStart('0');
            }

            if (trimmed.StartsWith("ZONE", StringComparison.OrdinalIgnoreCase))
            {
                var remainder = trimmed.Substring(4).Trim();
                return remainder.Length > 0 ? remainder : string.Empty;
            }

            return trimmed;
        }

        private static string GetBlockName(BlockReference br, Transaction tr)
        {
            if (br == null) return string.Empty;
            var btr = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
            return btr.Name;
        }

        private static Dictionary<string, string> ReadAttributes(BlockReference br, Transaction tr)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (br.AttributeCollection == null) return values;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                if (!attId.IsValid) continue;
                var attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (attRef == null) continue;
                values[attRef.Tag] = attRef.TextString;
            }

            return values;
        }

        private static string GetValue(IDictionary<string, string> dict, string key)
        {
            return (dict != null && dict.TryGetValue(key, out var value)) ? value : string.Empty;
        }

        private static string SanitizeAttributeValue(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var sanitized = value;

            sanitized = Regex.Replace(sanitized, @"\\S([^;]+);", m =>
            {
                var frac = m.Groups[1].Value.Replace('#', '/');
                return frac;
            }, RegexOptions.IgnoreCase);

            sanitized = Regex.Replace(sanitized, @"\\P|\\~", " ", RegexOptions.IgnoreCase);
            sanitized = Regex.Replace(sanitized, @"\\[A-Za-z][^;]*;", string.Empty);
            sanitized = sanitized.Replace("{", string.Empty).Replace("}", string.Empty);
            sanitized = sanitized.Replace('\u00A0', ' ');
            sanitized = Regex.Replace(sanitized, @"\s+", " ").Trim();

            return sanitized;
        }

        private static void WriteAttribute(Transaction tr, BlockReference br, string tag, string value)
        {
            if (br.AttributeCollection == null) return;

            foreach (ObjectId attId in br.AttributeCollection)
            {
                var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (attRef == null) continue;

                if (string.Equals(attRef.Tag, tag, StringComparison.OrdinalIgnoreCase))
                    attRef.TextString = value ?? string.Empty;
            }
        }

        public class ScanResult
        {
            public IList<CrossingRecord> Records { get; set; }
            public IDictionary<ObjectId, DuplicateResolver.InstanceContext> InstanceContexts { get; set; }
        }
    }
}


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\TableSync.cs
/////////////////////////////////////////////////////////////////////

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
            "XNUM",
            "XNUMBER",
            "NUMBER",
            "INDEX",
            "NO",
            "LABEL"
        };

        private const int MaxHeaderRowsToScan = 5;

        private static readonly Regex MTextFormattingCommandRegex = new Regex("\\\\[^;\\\\{}]*;", RegexOptions.Compiled);
        private static readonly Regex MTextResidualCommandRegex = new Regex("\\\\[^{}]", RegexOptions.Compiled);
        private static readonly Regex MTextSpecialCodeRegex = new Regex("%%[^\\s]+", RegexOptions.Compiled);

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
                                        UpdatePageTable(table, byKey, out matched, out updated);
                                        break;
                                    case XingTableType.LatLong:
                                        var rowIndexMap = BuildLatLongRowIndexMap(table.ObjectId, byKey?.Values);
                                        UpdateLatLongTable(table, byKey, out matched, out updated, rowIndexMap);
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

                    int matched = 0, updated = 0;

                    try
                    {
                        if (identifiedType == XingTableType.LatLong)
                        {
                            // Normal header-driven updater (program-generated tables)
                            UpdateLatLongTable(table, byKey, out matched, out updated, rowIndexMap);
                        }
                        else
                        {
                            // Try normal pass anyway
                            UpdateLatLongTable(table, byKey, out matched, out updated, rowIndexMap);

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
                .OrderBy(r => r, Comparer<CrossingRecord>.Create(CrossingRecord.CompareByCrossing))
                .ToList();

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
                        TrySetCellBlockScale(t, row, 0, 1.0);   // force block scale = 1.0
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

            var boldStyleId = EnsureBoldTextStyle(db, tr, "XING_BOLD", "Standard");
            var headerColor = Color.FromColorIndex(ColorMethod.ByAci, 254);
            var titleColor = Color.FromColorIndex(ColorMethod.ByAci, 14);

            if (showColumnHeaders)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    var cell = table.Cells[headerRow, c];
                    cell.TextHeight = CellTextHeight;
                    cell.Alignment = CellAlignment.MiddleCenter;
                    cell.TextStyleId = boldStyleId;
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
                titleCell.TextStyleId = boldStyleId;
                ApplyCellTextColor(titleCell, titleColor);
                ApplyTitleBorderStyle(table, titleRow, 0, table.Columns.Count - 1);
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
                    headerCell.TextStyleId = boldStyleId;
                    headerCell.TextString = headerText;
                    ApplyCellTextColor(headerCell, titleColor);
                    ApplyTitleBorderStyle(table, row, 0, table.Columns.Count - 1);
                    continue;
                }

                if (rowInfo.IsColumnHeader)
                {
                    ApplyLatLongColumnHeaderRow(table, row, HeaderRowHeight, CellTextHeight, boldStyleId, headerColor);
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
                SetBorderVisible(bordersObj, "Bottom", true);
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
                                if (string.Equals(text, "LATLONG", StringComparison.OrdinalIgnoreCase)) return XingTableType.LatLong;
                            }
                        }
                    }
                }
            }

            // Header-free heuristics: MAIN=5 cols, PAGE=3 cols, LATLONG=4 cols (+ sanity)
            if (table.Columns.Count == 5) return XingTableType.Main;
            if (table.Columns.Count == 3) return XingTableType.Page;

            if ((table.Columns.Count == 4 || table.Columns.Count >= 6) && table.Rows.Count >= 1)
            {
                var headerColumns = Math.Min(table.Columns.Count, 6);
                if (HasHeaderRow(table, headerColumns, IsLatLongHeader) || LooksLikeLatLongTable(table))
                    return XingTableType.LatLong;
            }

            return XingTableType.Unknown;
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

        private void UpdatePageTable(Table table, IDictionary<string, CrossingRecord> byKey, out int matched, out int updated)
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

                if (record == null) record = FindRecordByPageColumns(table, row, records);

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

                var rowUpdated = false;

                if (columnCount > 0 && ValueDiffers(rawKey, desiredCrossing)) rowUpdated = true;
                if (columnCount > 1 && ValueDiffers(ReadCellText(table, row, 1), desiredOwner)) rowUpdated = true;
                if (columnCount > 2 && ValueDiffers(ReadCellText(table, row, 2), desiredDescription)) rowUpdated = true;

                if (columnCount > 0) SetCellCrossingValue(table, row, 0, desiredCrossing);

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

            RefreshTable(table);
        }

        // TableSync.cs
        private void UpdateLatLongTable(
            Table table,
            IDictionary<string, CrossingRecord> byKey,
            out int matched,
            out int updated,
            IDictionary<int, CrossingRecord> rowIndexMap = null)
        {
            matched = 0;
            updated = 0;
            if (table == null) return;

            var columnCount = table.Columns.Count;
            var records = byKey?.Values ?? Enumerable.Empty<CrossingRecord>();
            var recordList = records.Where(r => r != null).ToList();

            var crossDescMap = BuildCrossingDescriptionMap(recordList);
            var descriptionMap = BuildDescriptionMap(recordList);

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

                // IMPORTANT:
                // - For LAT/LONG tables, the row identity is the cell contents (description + coordinates),
                //   NOT the bubble value. The bubble is only used to read the existing X# for comparison/update.
                var descriptionText = ReadCellTextSafe(table, row, 1);
                var descriptionKey = NormalizeDescriptionKey(descriptionText);

                CrossingRecord record = null;

                // 1) Prefer stable row->record mapping, but only if it still matches this row's description.
                if (rowIndexMap != null && rowIndexMap.TryGetValue(row, out var mapRecord))
                {
                    if (string.IsNullOrEmpty(descriptionKey) ||
                        string.Equals(NormalizeDescriptionKey(mapRecord.Description), descriptionKey, StringComparison.Ordinal))
                    {
                        record = mapRecord;
                    }
                }

                // 2) Primary match = table cell contents (description + lat/long [+ zone/dwg])
                if (record == null)
                {
                    record = FindRecordByLatLongColumns(table, row, recordList);
                }

                // 3) If description is unique, it is safe to match on that alone.
                if (record == null && !string.IsNullOrEmpty(descriptionKey))
                {
                    if (descriptionMap.TryGetValue(descriptionKey, out var list) && list.Count == 1)
                    {
                        record = list[0];
                    }
                }

                // 4) Last resort: use existing X# from the bubble, but ONLY if it matches the row description.
                if (record == null && !string.IsNullOrEmpty(key) && byKey.TryGetValue(key, out var keyed))
                {
                    if (string.IsNullOrEmpty(descriptionKey) ||
                        string.Equals(NormalizeDescriptionKey(keyed.Description), descriptionKey, StringComparison.Ordinal))
                    {
                        record = keyed;
                    }
                }

                if (record == null)
                {
                    // Logger.Debug in this project requires (Editor, message).
                    Logger.Debug(_ed, $"Table {table.ObjectId.Handle}: no match for lat/long row {row} (x='{rawKey}', desc='{descriptionText}')");
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

        private static bool IsCoordinateValue(string text, double min, double max)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            if (!double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var value)) return false;
            return value >= min && value <= max;
        }

        private static string ReadCellText(Cell cell)
        {
            if (cell == null) return string.Empty;
            try { return cell.TextString ?? string.Empty; } catch { return string.Empty; }
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
        // Ensures autoscale is OFF and final scale == 1.0
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
                // Force scale to 1 after placement (tableâ€‘level API if available, else perâ€‘content)
                TrySetCellBlockScale(table, row, col, 1.0);
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

            // 1) Try to set the block attribute by any of our known tags
            foreach (var tag in CrossingAttributeTags)
            {
                if (TrySetBlockAttributeValue(t, row, col, tag, crossingText))
                    return;
            }

            // 2) If the cell currently hosts a block, do NOT overwrite it with text.
            Cell cell = null;
            try { cell = t.Cells[row, col]; } catch { cell = null; }
            if (CellHasBlockContent(cell)) return;

            // 3) Fallback to plain text when no block
            try
            {
                if (cell != null) cell.TextString = crossingText ?? string.Empty;
            }
            catch { }
        }

        private static void SetCellValue(Cell cell, string value)
        {
            if (cell == null) return;

            var desired = value ?? string.Empty;

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
        // Try to set the scale of a block hosted in a table cell to a specific value (e.g., 1.0)
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
            if (candidates.Count == 1) return candidates[0];

            candidates = records.Where(r => string.Equals(Norm(r.Description), desc, StringComparison.Ordinal)).ToList();
            return candidates.Count == 1 ? candidates[0] : null;
        }

        private static CrossingRecord FindRecordByLatLongColumns(Table t, int row, IEnumerable<CrossingRecord> records)
        {
            var desc = ReadNorm(t, row, 1);
            var columns = t?.Columns.Count ?? 0;
            var zone = columns >= 6 ? ReadNorm(t, row, 2) : string.Empty;
            var lat = ReadNorm(t, row, columns >= 6 ? 3 : 2);
            var lng = ReadNorm(t, row, columns >= 6 ? 4 : 3);
            var dwg = columns >= 6 ? ReadNorm(t, row, 5) : string.Empty;

            var candidates = records.Where(r =>
                string.Equals(Norm(r.Description), desc, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Lat), lat, StringComparison.Ordinal) &&
                string.Equals(Norm(r.Long), lng, StringComparison.Ordinal) &&
                (string.IsNullOrEmpty(zone) || string.Equals(Norm(r.Zone), zone, StringComparison.Ordinal)) &&
                (string.IsNullOrEmpty(dwg) || string.Equals(Norm(r.DwgRef), dwg, StringComparison.Ordinal))).ToList();
            if (candidates.Count == 1) return candidates[0];

            candidates = records.Where(r => string.Equals(Norm(r.Description), desc, StringComparison.Ordinal)).ToList();
            return candidates.Count == 1 ? candidates[0] : null;
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

/////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\XingForm.cs
/////////////////////////////////////////////////////////////////////

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
            var nextIndex = 1;
            for (int i = 0; i < _records.Count; i++)
            {
                if (!IsExplicitXKey(_records[i].Crossing))
                    continue;

                string oldKey = NormalizeXKey(_records[i].Crossing);
                if (string.IsNullOrWhiteSpace(oldKey))
                    continue;

                string newKey = $"X{nextIndex}";
                nextIndex++;
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

            // Apply new X#s in list order for explicit X crossings only.
            nextIndex = 1;
            for (int i = 0; i < _records.Count; i++)
            {
                if (!IsExplicitXKey(_records[i].Crossing))
                    continue;

                _records[i].Crossing = $"X{nextIndex}";
                nextIndex++;
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

                            if (!IsExplicitXKey(raw))
                                continue;

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
        // - forces a real graphics refresh to avoid the ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œghost rowÃƒÂ¢Ã¢â€šÂ¬Ã‚Â
                
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
                        UpdateAllXingTablesFromGrid();               // safety XÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœkey pass (4/6ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœcol LAT/LONG)
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
                // Snapshot the grid BEFORE we touch the DWG; this is our ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œsource of truthÃƒÂ¢Ã¢â€šÂ¬Ã‚Â
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

                // 8) NEW: guard against ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œsnap backÃƒÂ¢Ã¢â€šÂ¬Ã‚Â only for LAT/LONG by reapplying the snapshotÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢s LAT/LONG.
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
        //      - Finds per-row LAT/LONG columns by label (ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œLATÃƒÂ¢Ã¢â€šÂ¬Ã‚Â, ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œLONGÃƒÂ¢Ã¢â€šÂ¬Ã‚Â, ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œLATITUDEÃƒÂ¢Ã¢â€šÂ¬Ã‚Â, ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œLONGITUDEÃƒÂ¢Ã¢â€šÂ¬Ã‚Â) or numeric range.
        //      - Writes only LAT and LONG to minimize side-effects.
        //      - Tags updated tables as LATLONG (so IdentifyTable will classify them next time).
        private void ReplaceLatLongInAnyTableRobust(IList<CrossingRecord> snapshot)
        {
            var doc = _doc ?? AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null || snapshot == null || snapshot.Count == 0) return;

            // Build a key -> record map (accept either ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œX12ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â or bare digits, normalize both ways)
            var byX = new Dictionary<string, CrossingRecord>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in snapshot)
            {
                if (r == null) continue;
                var k1 = TableSync.NormalizeKeyForLookup(r.Crossing);   // ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œX##ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â or ""  (robust normalization)  :contentReference[oaicite:10]{index=10}
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

                            // Force a redraw to avoid the ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œghost rowÃƒÂ¢Ã¢â€šÂ¬Ã‚Â look
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
        //      snapshotÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢s LAT/LONG (and Zone if you want) back into the grid so the form
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

            return latCol >= 0 && longCol >= 0 && latCol < cols && longCol < cols;
        }


        // --- NEW helper: only touches LAT, LONG (and ZONE/DWG_REF column in 6ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœcol LAT/LONG tables) ---
        private void BruteForceReplaceLatLongEverywhere(IList<CrossingRecord> snapshot)
        {
            var doc = _doc ?? AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null || snapshot == null || snapshot.Count == 0) return;

            // Build byÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬ËœX map from the current grid snapshot ("X3" etc.)
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

                        // Treat classic 4ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœ or 6ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœcolumn layouts as LAT/LONG candidates (ID|DESC|LAT|LONG) or (ID|DESC|ZONE|LAT|LONG|DWG_REF)
                        var cols = table.Columns.Count;
                        if (cols != 4 && cols != 6) continue;

                        // Use the same startÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœrow logic you already rely on elsewhere
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
                            // Column 0 key: attributeÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœfirst, fall back to visible token
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

                            // Write LAT/LONG (and Zone/DWG_REF if the 6ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœcol layout is present)
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

            // NonÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœundoable repaint paths
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
        /// Ends with a hard refresh to avoid the transient ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œghost rowÃƒÂ¢Ã¢â€šÂ¬Ã‚Â look.
        /// Update recognised crossing tables (MAIN/PAGE/LATLONG) from the current grid.
        /// Never modifies Column 0 (the bubble/X); only columns B..E are written.
        /// Ends with a hard refresh to avoid the transient ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œghost rowÃƒÂ¢Ã¢â€šÂ¬Ã‚Â look.
        /// Update recognised crossing tables (MAIN, PAGE, LATLONG) from the current grid.
        /// - Matches rows by X-key (attribute-first, with text fallback).
        /// - Never writes Column 0 (bubble) in any table.
        /// - For LAT/LONG, supports both 4ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœcol (ID, DESC, LAT, LONG) and 6ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœcol (ID, DESC, ZONE, LAT, LONG, DWG_REF).
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
                                // 4ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœcol: ID | DESC | LAT | LONG
                                // 6ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœcol: ID | DESC | ZONE | LAT | LONG | DWG_REF
                                var cols = table.Columns.Count;

                                // Description (col 1 is common)
                                changed |= SetCellIfChanged(table, row, 1, rec.Description);

                                // Zone label for 6ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœcol tables (e.g., "ZONE 11")
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

        // strip inline MTEXT formatting codes like \A1;, \H2.5x;, \C2;, {ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦}, etc.
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
                    CommandLogger.Log(ed, "--- TABLEÃƒÂ¢Ã¢â‚¬Â Ã¢â‚¬â„¢FORM UPDATES (X-only) ---");

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

                    // 2) Update MAIN/PAGE/LATLONG tables from the grid (B..E only ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Â never col 0)
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

            // Add any new records for rows that didnÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢t match existing crossings.
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

        private static bool IsExplicitXKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            var up = Regex.Replace(s.ToUpperInvariant(), @"\s+", "");
            return Regex.IsMatch(up, @"^X0*\d+$");
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
                            // MAIN/PAGE: if row 0 isnÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢t an X row, advance until we hit a row that looks like data.
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

                            // 2) Fallback: clear the row cells so we donÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢t leave stale data behind
                            if (!deleted && row >= 0 && row < table.Rows.Count)
                            {
                                try
                                {
                                    int cols = table.Columns.Count;
                                    for (int c = 0; c < cols; c++)
                                    {
                                        // Never touch Column 0ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢s block content; just blank the visible text.
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

        // Modify your btnDelete_Click handler so it doesnÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢t renumber bubbles
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

                // Mark graphics ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œdirtyÃƒÂ¢Ã¢â€šÂ¬Ã‚Â so the display list is rebuilt
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
                        if (!IsExplicitXKey(rec.Crossing))
                            continue;

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
                   string.Equals(trimmed, "ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Â", StringComparison.Ordinal) ||
                   string.Equals(trimmed, "ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Å“", StringComparison.Ordinal);
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

            // Build perÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬ËœDWG_REF list of distinct LOCATIONs
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

                        // Sort each sectionÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢s records numerically and sort sections by earliest crossing
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

                            // Insert each ownerÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢s other table separately
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
        /// (e.g., XING #1, XING #2, XING #3).  The ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œModelÃƒÂ¢Ã¢â€šÂ¬Ã‚Â tab (order 0)
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

            // Sort each sectionÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢s records by crossing number (ascending)
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

                // Sort each sectionÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢s records by crossing number
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
                // Check for the key phrase "CROSSING INFORMATION" (caseÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Ëœinsensitive)
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
                CommandLogger.Log(editor, "** Invalid zone ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Å“ enter 11 or 12. **", alsoToCommandBar: true);
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


/////////////////////////////////////////////////////////////////////
// FILE: C:\Users\Jesse 2025\Desktop\CROSSING MANAGER\XingManager\XingManagerApp.cs
/////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using XingManager.Services;

namespace XingManager
{
    public class XingManagerApp : IExtensionApplication
    {
        private static readonly Guid PaletteGuid = new Guid("71E8DF88-8F04-4D7E-AD5F-97F1F4F0F5BB");

        private PaletteSet _palette;
        private XingForm _form;
        private Document _formDocument;
        private readonly TableFactory _tableFactory = new TableFactory();
        private readonly LayoutUtils _layoutUtils = new LayoutUtils();
        private readonly Serde _serde = new Serde();
        private readonly DuplicateResolver _duplicateResolver = new DuplicateResolver();
        private readonly LatLongDuplicateResolver _latLongDuplicateResolver = new LatLongDuplicateResolver();
        private TableSync _tableSync;

        public static XingManagerApp Instance { get; private set; }

        public void Initialize()
        {
            Instance = this;
            _tableSync = new TableSync(_tableFactory);
            Application.DocumentManager.DocumentActivated += OnDocumentActivated;
        }

        public void Terminate()
        {
            Application.DocumentManager.DocumentActivated -= OnDocumentActivated;
            if (_palette != null)
            {
                _palette.Visible = false;
                if (_palette.Count > 0)
                {
                    _palette.Remove(0);
                }

                _form?.Dispose();
                _palette.Dispose();
                _palette = null;
            }

            _form = null;
            _formDocument = null;
            Instance = null;
        }

        private void OnDocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            if (_palette == null || !_palette.Visible)
            {
                return;
            }

            _form = null;
            ShowPalette();
        }

        internal void ShowPalette()
        {
            var form = GetOrCreateForm();
            if (form == null)
            {
                return;
            }

            EnsurePalette();
            ActivatePalette();
        }

        internal XingForm GetOrCreateForm()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return null;
            }

            if (_form == null || _formDocument != doc)
            {
                _form = CreateForm(doc);
                _formDocument = doc;
            }

            return _form;
        }

        private XingForm CreateForm(Document doc)
        {
            var repository = new XingRepository(doc);
            var form = new XingForm(
                doc,
                repository,
                _tableSync,
                _layoutUtils,
                _tableFactory,
                _serde,
                _duplicateResolver,
                _latLongDuplicateResolver);
            form.LoadData();
            AttachForm(form);
            return form;
        }

        private void AttachForm(XingForm form)
        {
            EnsurePalette();
            if (_palette.Count > 0)
            {
                _palette.Remove(0);
                _form?.Dispose();
            }

            _palette.Add("Crossings", form);
            _palette.Visible = true;
            if (_palette.Count > 0)
            {
                _palette.Activate(0);
            }
        }

        private void EnsurePalette()
        {
            if (_palette != null)
            {
                return;
            }

            _palette = new PaletteSet("Crossing Manager", PaletteGuid)
            {
                Style = PaletteSetStyles.ShowAutoHideButton | PaletteSetStyles.ShowCloseButton,
                MinimumSize = new Size(500, 400)
            };
        }

        private void ActivatePalette()
        {
            if (_palette == null)
            {
                return;
            }

            _palette.Visible = true;
            if (_palette.Count > 0)
            {
                _palette.Activate(0);
            }
        }
    }
}
