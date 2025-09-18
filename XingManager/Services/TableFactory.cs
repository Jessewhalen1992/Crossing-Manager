using System;
using System.Collections.Generic;
using System.Linq;
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

            var table = new Table
            {
                TableStyle = EnsureTableStyle(db, tr),
                LayerId = LayerUtils.EnsureLayer(db, tr, LayerName)
            };

            var rows = recordList.Count + 1;
            table.NumRows = rows;
            table.NumColumns = 5;

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
            return table;
        }

        public Table CreateCrossingPageTable(Database db, Transaction tr, string dwgRef, IEnumerable<CrossingRecord> records)
        {
            var filtered = PrepareRecordList(records)
                .Where(r => string.Equals((r.DwgRef ?? string.Empty).Trim(), (dwgRef ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase))
                .ToList();

            var table = new Table
            {
                TableStyle = EnsureTableStyle(db, tr),
                LayerId = LayerUtils.EnsureLayer(db, tr, LayerName)
            };

            var rows = filtered.Count + 1;
            table.NumRows = Math.Max(1, rows);
            table.NumColumns = 3;
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
            return table;
        }

        public Table CreateLatLongTable(Database db, Transaction tr, CrossingRecord record)
        {
            if (record == null)
            {
                throw new ArgumentNullException("record");
            }

            var table = new Table
            {
                TableStyle = EnsureTableStyle(db, tr),
                LayerId = LayerUtils.EnsureLayer(db, tr, LayerName)
            };

            table.NumRows = 2;
            table.NumColumns = 4;
            table.SetRowHeight(25.0);
            table.Columns[0].Width = 40.0;
            table.Columns[1].Width = 150.0;
            table.Columns[2].Width = 90.0;
            table.Columns[3].Width = 90.0;

            var textStyleId = db.Textstyle;
            const double textHeight = 10.0;
            var headers = new[] { "XING", "DESCRIPTION", "LAT", "LONG" };
            for (var col = 0; col < headers.Length; col++)
            {
                SetCellText(table, 0, col, headers[col], textHeight, textStyleId);
            }

            SetCellText(table, 1, 0, record.Crossing, textHeight, textStyleId);
            SetCellText(table, 1, 1, record.Description, textHeight, textStyleId);
            SetCellText(table, 1, 2, record.Lat, textHeight, textStyleId);
            SetCellText(table, 1, 3, record.Long, textHeight, textStyleId);

            table.GenerateLayout();
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
