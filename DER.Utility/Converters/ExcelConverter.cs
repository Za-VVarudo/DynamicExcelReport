using ClosedXML.Excel;
using DER.Utility.Helpers;
using DER.Utility.Models;
using Newtonsoft.Json;

namespace DER.Utility.Converters
{
    public partial class ExcelConverter
    {
        private static readonly JsonSerializerSettings DefaultSerializerSetting = new()
        {
            Formatting = Formatting.None,
            NullValueHandling = NullValueHandling.Ignore
        };
        private const XLCellsUsedOptions DefaultCellsUsedOptions = XLCellsUsedOptions.AllContents | XLCellsUsedOptions.MergedRanges | XLCellsUsedOptions.NormalFormats;

        public static string ExcelToJson(FileStream stream)
        {
            var workbookData = new WorkbookModel();
            using (var workbook = new XLWorkbook(stream))
            {
                var themes = workbook.Theme;
                foreach (var ws in workbook.Worksheets)
                {
                    var table = ws.Tables;
                    var sheetData = new SheetModel
                    {
                        Cells = [],
                        ColumnWidths = [],
                        RowHeights = []
                    };
                    HashSet<string> mergedRanges = [];
                    foreach (var cell in ws.CellsUsed(DefaultCellsUsedOptions))
                    {
                        var mergedRange = cell.MergedRange()?.RangeAddress.ToString();
                        if (!string.IsNullOrEmpty(mergedRange) && !mergedRanges.Add(mergedRange))
                        {
                            continue;
                        }
                        var column = cell.WorksheetColumn();
                        var row = cell.WorksheetRow();
                        sheetData.ColumnWidths.TryAdd(column.ColumnNumber(), column.Width);
                        sheetData.RowHeights.TryAdd(row.RowNumber(), row.Height);

                        var cellAddress = cell.Address.ToString();
                        var value = cell.GetFormattedString();

                        // Style extraction
                        var style = cell.Style;
                        var border = style.Border;

                        var cellInfo = new CellDataModel
                        {
                            Value = value,
                            MergedRange = mergedRange,
                            FontName = style.Font.FontName,
                            FontSize = style.Font.FontSize,
                            Bold = style.Font.Bold,
                            Italic = style.Font.Italic,
                            FontColor = ToArgbHex(style.Font.FontColor, themes),
                            FillForegroundColor = ToArgbHex(style.Fill.PatternColor, themes),
                            FillBackgroundColor = ToArgbHex(style.Fill.BackgroundColor, themes),
                            PatternType = (int)style.Fill.PatternType,
                            HorizontalAlign = style.Alignment.Horizontal.ToString(),
                            VerticalAlign = style.Alignment.Vertical.ToString(),

                            // Borders: set null if None, otherwise cast to int
                            BorderTopStyle = border.TopBorder != XLBorderStyleValues.None ? (int)border.TopBorder : null,
                            BorderTopColor = ToArgbHex(border.TopBorderColor, themes),
                            BorderBottomStyle = border.BottomBorder != XLBorderStyleValues.None ? (int)border.BottomBorder : null,
                            BorderBottomColor = ToArgbHex(border.BottomBorderColor, themes),
                            BorderLeftStyle = border.LeftBorder != XLBorderStyleValues.None ? (int)border.LeftBorder : null,
                            BorderLeftColor = ToArgbHex(border.LeftBorderColor, themes),
                            BorderRightStyle = border.RightBorder != XLBorderStyleValues.None ? (int)border.RightBorder : null,
                            BorderRightColor = ToArgbHex(border.RightBorderColor, themes),
                            BorderDiagonalStyle = border.DiagonalBorder != XLBorderStyleValues.None ? (int)border.DiagonalBorder : null,
                            BorderDiagonalColor = ToArgbHex(border.DiagonalBorderColor, themes),
                            BorderDiagonalUp = border.DiagonalUp ? border.DiagonalUp : null,
                            BorderDiagonalDown = border.DiagonalDown ? border.DiagonalDown : null
                        };

                        var hasFill = !string.IsNullOrEmpty(cellInfo.FillForegroundColor) || !string.IsNullOrEmpty(cellInfo.FillBackgroundColor);
                        var hasBorders = cellInfo.BorderTopStyle.HasValue || cellInfo.BorderBottomStyle.HasValue || cellInfo.BorderLeftStyle.HasValue || cellInfo.BorderRightStyle.HasValue || cellInfo.BorderDiagonalStyle.HasValue || (cellInfo.BorderDiagonalUp == true) || (cellInfo.BorderDiagonalDown == true);

                        if (!string.IsNullOrEmpty(value) || !string.IsNullOrEmpty(cellInfo.MergedRange) || hasFill || hasBorders)
                        {
                            sheetData.Cells[cellAddress] = cellInfo;
                        }
                    }
                    workbookData[ws.Name] = sheetData;
                }
            }

            return JsonConvert.SerializeObject(workbookData, DefaultSerializerSetting);
        }

        public static void JsonToExcel(string jsonString, string outputPath)
        {
            var excelTemplate = JsonConvert.DeserializeObject<WorkbookModel>(jsonString) ?? throw new ArgumentException("Invalid JSON format");

            using var workbook = new XLWorkbook();

            foreach (var sheetData in excelTemplate)
            {
                var ws = workbook.Worksheets.Add(sheetData.Key);
                var mergesToApply = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var cellData in sheetData.Value.Cells)
                {
                    var cellAddress = cellData.Key;
                    var cellInfo = cellData.Value;

                    var cell = ws.Cell(cellAddress);

                    // Set cell value
                    if (!string.IsNullOrEmpty(cellInfo.Value))
                    {
                        cell.Value = cellInfo.Value;
                    }

                    // Apply styling using flattened model
                    ApplyStyle(cell.Style, cellInfo);

                    // Handle merged cells (only add once per range)
                    if (!string.IsNullOrEmpty(cellInfo.MergedRange))
                    {
                        if (mergesToApply.Add(cellInfo.MergedRange))
                        {
                            var range = ws.Range(cellInfo.MergedRange);
                            ApplyStyle(range.Style, cellInfo);
                            range.Merge();
                        }
                    }
                }

                foreach (var columnWidth in sheetData.Value.ColumnWidths)
                {
                    ws.Column(columnWidth.Key).Width = columnWidth.Value;
                }
                foreach (var rowHeight in sheetData.Value.RowHeights)
                {
                    ws.Row(rowHeight.Key).Height = rowHeight.Value;
                }
            }

            workbook.SaveAs(outputPath);
        }

        private static void ApplyStyle(IXLStyle cellStyle, CellDataModel style)
        {
            if (!string.IsNullOrEmpty(style.FontName)) cellStyle.Font.FontName = style.FontName;
            if (style.FontSize > 0) cellStyle.Font.FontSize = style.FontSize;
            cellStyle.Font.Bold = style.Bold;
            cellStyle.Font.Italic = style.Italic;
            if (!string.IsNullOrEmpty(style.FontColor))
            {
                cellStyle.Font.FontColor = XLColor.FromHtml(style.FontColor);
            }
            cellStyle.Fill.PatternType = (XLFillPatternValues)style.PatternType;
            if (!string.IsNullOrEmpty(style.FillBackgroundColor))
            {
                cellStyle.Fill.BackgroundColor = XLColor.FromHtml(style.FillBackgroundColor);
            }
            if (!string.IsNullOrEmpty(style.FillForegroundColor))
            {
                cellStyle.Fill.PatternColor = XLColor.FromHtml(style.FillForegroundColor);
            }

            if (!string.IsNullOrEmpty(style.HorizontalAlign))
            {
                if (Enum.TryParse<XLAlignmentHorizontalValues>(style.HorizontalAlign, true, out var h))
                {
                    cellStyle.Alignment.Horizontal = h;
                }
            }
            if (!string.IsNullOrEmpty(style.VerticalAlign))
            {
                if (Enum.TryParse<XLAlignmentVerticalValues>(style.VerticalAlign, true, out var v))
                {
                    cellStyle.Alignment.Vertical = v;
                }
            }

            // Borders per side
            if (style.BorderTopStyle.HasValue)
            {
                cellStyle.Border.TopBorder = (XLBorderStyleValues)style.BorderTopStyle.Value;
                if (!string.IsNullOrEmpty(style.BorderTopColor))
                {
                    cellStyle.Border.TopBorderColor = XLColor.FromHtml(style.BorderTopColor);
                }
            }
            if (style.BorderBottomStyle.HasValue)
            {
                cellStyle.Border.BottomBorder = (XLBorderStyleValues)style.BorderBottomStyle.Value;
                if (!string.IsNullOrEmpty(style.BorderBottomColor))
                {
                    cellStyle.Border.BottomBorderColor = XLColor.FromHtml(style.BorderBottomColor);
                }
            }
            if (style.BorderLeftStyle.HasValue)
            {
                cellStyle.Border.LeftBorder = (XLBorderStyleValues)style.BorderLeftStyle.Value;
                if (!string.IsNullOrEmpty(style.BorderLeftColor))
                {
                    cellStyle.Border.LeftBorderColor = XLColor.FromHtml(style.BorderLeftColor);
                }
            }
            if (style.BorderRightStyle.HasValue)
            {
                cellStyle.Border.RightBorder = (XLBorderStyleValues)style.BorderRightStyle.Value;
                if (!string.IsNullOrEmpty(style.BorderRightColor))
                {
                    cellStyle.Border.RightBorderColor = XLColor.FromHtml(style.BorderRightColor);
                }
            }
            if (style.BorderDiagonalStyle.HasValue)
            {
                cellStyle.Border.DiagonalUp = style.BorderDiagonalUp ?? false;
                cellStyle.Border.DiagonalDown = style.BorderDiagonalDown ?? false;
                cellStyle.Border.DiagonalBorder = (XLBorderStyleValues)style.BorderDiagonalStyle.Value;
                if (!string.IsNullOrEmpty(style.BorderDiagonalColor))
                {
                    cellStyle.Border.DiagonalBorderColor = XLColor.FromHtml(style.BorderDiagonalColor);
                }
            }
        }

        private static string ToArgbHex(XLColor color, IXLTheme themes)
        {
            if (color == null) return null;
            double tint = 0;
            switch (color.ColorType)
            {
                case XLColorType.Color:
                    break;
                case XLColorType.Indexed:
                    color = XLColor.FromIndex(color.Indexed);
                    break;
                case XLColorType.Theme:
                    tint = color.ThemeTint;
                    color = themes.ResolveThemeColor(color.ThemeColor);
                    break;
                default:
                    return null;
            }
            return ColorHelpers.ApplyTint(color.Color, tint); ;
        }
    }
}
