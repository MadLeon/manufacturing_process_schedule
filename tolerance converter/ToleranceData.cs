using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ToleranceConverter
{
    public enum ToleranceType
    {
        Internal,
        External,
        IT12Half
    }

    public class ToleranceRange
    {
        [JsonPropertyName("minRange")]
        public double MinRange { get; set; }

        [JsonPropertyName("maxRange")]
        public double MaxRange { get; set; }

        [JsonPropertyName("upper")]
        public double Upper { get; set; }

        [JsonPropertyName("lower")]
        public double Lower { get; set; }
    }

    public class ToleranceTable
    {
        [JsonPropertyName("internal")]
        public List<ToleranceRange> Internal { get; set; } = new List<ToleranceRange>();

        [JsonPropertyName("external")]
        public List<ToleranceRange> External { get; set; } = new List<ToleranceRange>();

        [JsonPropertyName("it12half")]
        public List<ToleranceRange> It12Half { get; set; } = new List<ToleranceRange>();
    }

    public class ToleranceDataService
    {
        private ToleranceTable? _toleranceTable;

        public ToleranceDataService()
        {
            LoadDataFromEmbeddedResource();
        }

        private void LoadDataFromEmbeddedResource()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var resourceName = "ToleranceConverter.tolerance_table.json";

                using (Stream? stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                        throw new Exception($"Could not find embedded resource: {resourceName}");

                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string jsonContent = reader.ReadToEnd();
                        _toleranceTable = JsonSerializer.Deserialize<ToleranceTable>(jsonContent);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to load tolerance data from embedded resource: {ex.Message}");
            }
        }

        /// <summary>
        /// Returns upper and lower tolerance bounds for the given dimension and tolerance type.
        /// Range lookup uses exclusive min, inclusive max: (minRange, maxRange].
        /// </summary>
        /// <param name="dimension">Nominal dimension in millimeters</param>
        /// <param name="type">Tolerance type: Internal (H12), External (h12), or IT12Half</param>
        /// <returns>Tuple of (upper, lower) tolerance values in mm, or null if not found</returns>
        public (double upper, double lower)? GetTolerance(double dimension, ToleranceType type)
        {
            if (_toleranceTable == null)
                return null;

            List<ToleranceRange> ranges = type switch
            {
                ToleranceType.Internal => _toleranceTable.Internal,
                ToleranceType.External => _toleranceTable.External,
                ToleranceType.IT12Half => _toleranceTable.It12Half,
                _ => _toleranceTable.Internal
            };

            var match = ranges.FirstOrDefault(r => dimension > r.MinRange && dimension <= r.MaxRange);
            return match != null ? (match.Upper, match.Lower) : null;
        }
    }
}
