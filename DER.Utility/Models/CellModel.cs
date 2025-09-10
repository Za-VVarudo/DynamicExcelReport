using Newtonsoft.Json;

namespace DER.Utility.Models
{
    public class WorkbookModel : Dictionary<string, SheetModel> { }
    public class SheetModel {
        [JsonProperty("columnWidths")]
        public Dictionary<int, double> ColumnWidths { get; set; }
        [JsonProperty("rowHeights")]
        public Dictionary<int, double> RowHeights { get; set; }
        [JsonProperty("cells")]
        public Dictionary<string, CellDataModel> Cells { get; set; }
    }

    public class CellDataModel
    {
        [JsonProperty("value")]
        public string Value { get; set; }

        [JsonProperty("mergedRange")]
        public string MergedRange { get; set; }

        // Flattened style properties
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

        [JsonProperty("fillPatternType")]
        public int PatternType { get; set; }

        [JsonProperty("horizontalAlign")]
        public string HorizontalAlign { get; set; }

        [JsonProperty("verticalAlign")]
        public string VerticalAlign { get; set; }

        [JsonProperty("borderTopStyle")]
        public int? BorderTopStyle { get; set; }

        [JsonProperty("borderTopColor")]
        public string BorderTopColor { get; set; }

        [JsonProperty("borderBottomStyle")]
        public int? BorderBottomStyle { get; set; }

        [JsonProperty("borderBottomColor")]
        public string BorderBottomColor { get; set; }

        [JsonProperty("borderLeftStyle")]
        public int? BorderLeftStyle { get; set; }

        [JsonProperty("borderLeftColor")]
        public string BorderLeftColor { get; set; }

        [JsonProperty("borderRightStyle")]
        public int? BorderRightStyle { get; set; }

        [JsonProperty("borderRightColor")]
        public string BorderRightColor { get; set; }

        [JsonProperty("borderDiagonalStyle")]
        public int? BorderDiagonalStyle { get; set; }

        [JsonProperty("borderDiagonalColor")]
        public string BorderDiagonalColor { get; set; }

        [JsonProperty("borderDiagonalUp")]
        public bool? BorderDiagonalUp { get; set; }

        [JsonProperty("borderDiagonalDown")]
        public bool? BorderDiagonalDown { get; set; }
    }
}
