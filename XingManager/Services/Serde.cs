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
