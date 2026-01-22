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

