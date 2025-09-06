using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;

namespace DER.Utility.Converters
{
    public class ExcelConverter
    {
        public static string ExcelToJson(FileStream stream)
        {
            var workbookData = new Dictionary<string, Dictionary<string, object>>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
            {
                var workbookPart = document.WorkbookPart;
                var themePart = workbookPart.ThemePart;
                var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
                var stylesPart = workbookPart.WorkbookStylesPart;
                var stylesheet = stylesPart.Stylesheet;

                var fonts = stylesheet.Fonts.Elements<Font>().ToList();
                var fills = stylesheet.Fills.Elements<Fill>().ToList();
                var cellFormats = stylesheet.CellFormats.Elements<CellFormat>().ToList();

                foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    Worksheet worksheet = worksheetPart.Worksheet;
                    var sheetData = new Dictionary<string, object>();

                    var mergedCells = worksheet.Elements<MergeCells>().FirstOrDefault()?.Elements<MergeCell>()
                        .ToDictionary(mc => mc.Reference.Value.Split(':')[0], mc => mc.Reference.Value);

                    foreach (Row row in worksheet.Descendants<Row>())
                    {
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            string cellValue = GetCellValue(cell, sharedStrings);
                            string cellAddress = cell.CellReference;
                            string mergedRange = mergedCells != null && mergedCells.TryGetValue(cellAddress, out string value) ? value : string.Empty;
                            string fillForegroundColor = string.Empty;
                            string fillBackgroundColor = string.Empty;

                            var styleIndex = cell.StyleIndex;
                            var styleInfo = new Dictionary<string, object>();

                            if (styleIndex?.HasValue == true && styleIndex.Value < cellFormats.Count)
                            {
                                var cellFormat = cellFormats[(int)styleIndex.Value];

                                // Font
                                var font = fonts[(int)cellFormat.FontId.Value];
                                styleInfo["fontName"] = font.FontName?.Val?.Value;
                                styleInfo["fontSize"] = font.FontSize?.Val?.Value;
                                styleInfo["bold"] = font.Bold != null;
                                styleInfo["italic"] = font.Italic != null;
                                styleInfo["fontColor"] = font.Color?.Rgb?.Value ?? "";

                                // Fill
                                var fill = fills[(int)cellFormat.FillId.Value];
                                var patternFill = fill.PatternFill;
                                styleInfo["fillForegroundColor"] = fillForegroundColor = GetCellFillValue(themePart, stylesPart, patternFill?.ForegroundColor);
                                styleInfo["fillBackgroundColor"] = fillBackgroundColor = GetCellFillValue(themePart, stylesPart, patternFill?.BackgroundColor);

                                // Alignment
                                var alignment = cellFormat.Alignment;
                                styleInfo["horizontalAlign"] = alignment?.Horizontal?.InnerText.ToString() ?? "";
                                styleInfo["verticalAlign"] = alignment?.Vertical?.InnerText.ToString() ?? "";
                            }
                            else
                            {
                                Console.WriteLine($"Cell {cellAddress} does not have style");
                            }

                            var cellInfo = new Dictionary<string, object>
                            {
                                ["value"] = cellValue,
                                ["mergedRange"] = mergedRange,
                                ["style"] = styleInfo
                            };

                            if (!string.IsNullOrEmpty(cellValue) || !string.IsNullOrEmpty(mergedRange) ||
                                !string.IsNullOrEmpty(fillForegroundColor) || !string.IsNullOrEmpty(fillBackgroundColor))
                            {
                                sheetData[cellAddress] = cellInfo;
                            }
                        }
                    }

                    workbookData[sheet.Name] = sheetData;
                }
            }
            return JsonConvert.SerializeObject(workbookData, Formatting.None);
        }

        private static string GetCellValue(Cell cell, SharedStringTable sharedStrings)
        {
            if (cell == null || cell.CellValue == null) return string.Empty;
            var value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return sharedStrings?.ElementAt(int.Parse(value)).InnerText ?? value;
            }

            return value;
        }

        private static string GetCellFillValue(ThemePart themePart, WorkbookStylesPart stylesPart, ColorType colorType)
        {
            string value = string.Empty;
            switch (true)
            {
                case bool when colorType == null:
                    return value;

                case bool when colorType.Rgb != null:
                    value = colorType.Rgb?.Value;
                    break;

                case bool when colorType.Indexed != null && stylesPart.Stylesheet.Colors.IndexedColors != null:
                    var indexedColors = (IndexedColors)stylesPart.Stylesheet.Colors.IndexedColors.ChildElements[(int)colorType.Indexed.Value];
                    value = indexedColors.InnerText;
                    break;

                case bool when colorType.Theme != null:
                    var themeColor = (DocumentFormat.OpenXml.Drawing.Color2Type)themePart.Theme.ThemeElements.ColorScheme.ChildElements[(int)colorType.Theme.Value];
                    value = themeColor.RgbColorModelHex?.Val;
                    break;
            }
            return value ?? string.Empty;
        }
    }
}
