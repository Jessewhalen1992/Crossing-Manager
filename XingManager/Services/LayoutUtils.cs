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


        // Variant that also treats MTEXT non‑breaking spaces (\~) as whitespace
        private static readonly string Ws = @"(?:\s+|\\~)+";
        private static readonly Regex PlanHeadingRegexMText = new Regex(
            $"PLAN{Ws}SHOWING{Ws}PIPELINE{Ws}CROSSING\\(S\\){Ws}WITHIN(?:{Ws}AND{Ws}ADJACENT{Ws}TO)?",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static bool TryUpdateHeadingInMText(ref string contents, string replacement)
        {
            if (string.IsNullOrEmpty(contents)) return false;

            // 1) Try straight replace in the raw contents (keeps leading \H, \f, etc.)
            var updated = PlanHeadingRegex.Replace(contents, replacement);

            // 2) If not found, try the MTEXT‑aware variant that understands \~
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

