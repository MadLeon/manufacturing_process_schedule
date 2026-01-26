using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ToleranceConverter
{
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
                    {
                        throw new Exception($"Could not find embedded resource: {resourceName}");
                    }
                    
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

        public (double upper, double lower)? GetTolerance(double dimension, bool isInternal)
        {
            if (_toleranceTable == null)
                return null;

            var ranges = isInternal ? _toleranceTable.Internal : _toleranceTable.External;

            var matchedRange = ranges.FirstOrDefault(r => 
                dimension > r.MinRange && dimension <= r.MaxRange);

            if (matchedRange != null)
            {
                return (matchedRange.Upper, matchedRange.Lower);
            }

            return null;
        }
    }
}
