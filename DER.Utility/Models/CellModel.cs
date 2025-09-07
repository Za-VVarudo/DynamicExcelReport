using Newtonsoft.Json;

namespace DER.Utility.Models
{
    public class WorkbookModel : Dictionary<string, SheetModel> { }
    public class SheetModel : Dictionary<string, CellDataModel> { }

    public class CellDataModel
    {
        [JsonProperty("value")]
        public string Value { get; set; }

        [JsonProperty("mergedRange")]
        public string MergedRange { get; set; }

        [JsonProperty("style")]
        public CellStyleModel Style { get; set; }
    }

    public class CellStyleModel
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
