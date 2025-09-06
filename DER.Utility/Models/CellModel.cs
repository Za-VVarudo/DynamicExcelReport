using Newtonsoft.Json;

namespace DER.Utility.Models
{
    public class ExcelTemplate
    {
        [JsonProperty("Status Report")]
        public Dictionary<string, CellData> StatusReport { get; set; }
    }

    public class CellData
    {
        [JsonProperty("value")]
        public string Value { get; set; }

        [JsonProperty("mergedRange")]
        public string MergedRange { get; set; }

        [JsonProperty("style")]
        public CellStyle Style { get; set; }
    }

    public class CellStyle
    {
        [JsonProperty("fontName")]
        public string FontName { get; set; }

        [JsonProperty("fontSize")]
        public double FontSize { get; set; }

        [JsonProperty("bold")]
        public bool Bold { get; set; }

        [JsonProperty("italic")]
        public bool Italic { get; set; }

        [JsonProperty("fontColor")]
        public string FontColor { get; set; }

        [JsonProperty("fillForegroundColor")]
        public string FillForegroundColor { get; set; }

        [JsonProperty("fillBackgroundColor")]
        public string FillBackgroundColor { get; set; }

        [JsonProperty("horizontalAlign")]
        public string HorizontalAlign { get; set; }

        [JsonProperty("verticalAlign")]
        public string VerticalAlign { get; set; }
    }
}
