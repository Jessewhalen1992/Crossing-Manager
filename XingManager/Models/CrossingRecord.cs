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

        public List<ObjectId> AllInstances { get; } = new List<ObjectId>();

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
    }
}
