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
                doc.Editor.WriteMessage("\nUnable to switch layout: {0}", ex.Message);
            }
        }

        public void ReplacePlaceholderText(Database db, ObjectId layoutId, string replacement)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                foreach (ObjectId entId in btr)
                {
                    var dbText = tr.GetObject(entId, OpenMode.ForRead) as DBText;
                    if (dbText != null)
                    {
                        var text = dbText.TextString ?? string.Empty;
                        if (IsPlaceholder(text))
                        {
                            dbText.UpgradeOpen();
                            dbText.TextString = replacement;
                        }

                        continue;
                    }

                    var mtext = tr.GetObject(entId, OpenMode.ForRead) as MText;
                    if (mtext != null)
                    {
                        var contents = mtext.Text ?? mtext.Contents;
                        if (IsPlaceholder(contents))
                        {
                            mtext.UpgradeOpen();
                            mtext.Contents = replacement;
                        }
                    }
                }

                tr.Commit();
            }
        }

        private static bool IsPlaceholder(string value)
        {
            return string.Equals((value ?? string.Empty).Trim(), LocationPlaceholder, StringComparison.OrdinalIgnoreCase);
        }

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
