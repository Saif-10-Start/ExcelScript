using System.Text.Json.Serialization;

namespace ExcelScript
{
    [JsonSerializable(typeof(Config))]
    internal partial class AppJsonContext : JsonSerializerContext
    {
    }

    public class Config
    {
        [JsonPropertyName("headerColor")]
        public string HeaderColor { get; set; } = "#95b3d7";

        [JsonPropertyName("alternatingColors")]
        public string[] AlternatingColors { get; set; } = ["#c5d9f1", "#FFFFFF"];

        [JsonPropertyName("headerFontSize")]
        public int HeaderFontSize { get; set; } = 14;

        [JsonPropertyName("bodyFontSize")]
        public int BodyFontSize { get; set; } = 12;

        [JsonPropertyName("enableGrouping")]
        public bool GroupingEnabled { get; set; } = true;

        [JsonPropertyName("generateSummary")]
        public bool GenerateSummary { get; set; } = true;

        [JsonPropertyName("copyComments")]
        public bool CopyComments { get; set; } = true;

        [JsonPropertyName("enableStyling")]
        public bool StylingEnabled { get; set; } = true;
    }
}