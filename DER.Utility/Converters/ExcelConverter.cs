using DER.Utility.Constants;
using DER.Utility.Extensions;
using DER.Utility.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;

namespace DER.Utility.Converters
{
    public partial class ExcelConverter
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
                            string fillForegroundColor = null;
                            string fillBackgroundColor = null;

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
                                fillForegroundColor != null || fillBackgroundColor != null)
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
            if (colorType == null) return null;
            if (colorType.Rgb != null) return colorType.Rgb?.Value;
            if (colorType.Theme != null)
            {
                var themeColor = (DocumentFormat.OpenXml.Drawing.Color2Type)themePart.Theme.ThemeElements.ColorScheme.ChildElements[(int)colorType.Theme.Value];
                return themeColor.RgbColorModelHex.ToRGBA();
            }
            if (colorType.Indexed != null)
            {
                if (stylesPart.Stylesheet.Colors.IndexedColors != null)
                {
                    var indexedColors = (IndexedColors)stylesPart.Stylesheet.Colors.IndexedColors.ChildElements[(int)colorType.Indexed.Value];
                    return indexedColors.InnerText;
                }
                else if (colorType.Indexed.Value < ColorConstants.DefaultIndexedColors.Count)
                {
                    return ColorConstants.DefaultIndexedColors[(int)colorType.Indexed.Value];
                }
            }
            return null;
        }

        public static void JsonToExcel(string jsonString, string outputPath)
        {
            // Parse JSON to ExcelTemplate
            var excelTemplate = JsonConvert.DeserializeObject<WorkbookModel>(jsonString) ?? throw new ArgumentException("Invalid JSON format");
            using SpreadsheetDocument document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook);
            // Create workbook and worksheet parts
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add styles part
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = CreateStylesheet();

            // Add shared strings part
            var sharedStringsPart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringsPart.SharedStringTable = new SharedStringTable();

            // Add sheets
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());

            foreach (var sheetData in excelTemplate)
            {
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var worksheet = new Worksheet();
                var sheetDataElement = new SheetData();

                // Process cells and create rows
                var rows = new Dictionary<uint, Row>();
                var mergedCells = new List<MergeCell>();

                foreach (var cellData in sheetData.Value)
                {
                    var cellAddress = cellData.Key;
                    var cellInfo = cellData.Value;

                    // Parse cell address to get row and column
                    var (rowIndex, columnIndex) = ParseCellAddress(cellAddress);

                    // Get or create row
                    if (!rows.TryGetValue(rowIndex, out Row row))
                    {
                        row = new Row() { RowIndex = rowIndex };
                        rows[rowIndex] = row;
                    }

                    // Create cell
                    var cell = new Cell()
                    {
                        CellReference = cellAddress,
                        DataType = CellValues.String,
                        CellValue = new CellValue(string.Empty)
                    };

                    // Set cell value
                    if (!string.IsNullOrEmpty(cellInfo.Value))
                    {
                        cell.CellValue = new CellValue(cellInfo.Value);
                    }

                    // Apply styling
                    if (cellInfo.Style != null)
                    {
                        var styleIndex = GetOrCreateStyleIndex(stylesPart.Stylesheet, cellInfo.Style);
                        cell.StyleIndex = styleIndex;
                    }

                    // Handle merged cells
                    if (!string.IsNullOrEmpty(cellInfo.MergedRange))
                    {
                        mergedCells.Add(new MergeCell() { Reference = cellInfo.MergedRange });
                    }

                    row.AppendChild(cell);
                }

                // Add rows to sheet data
                foreach (var row in rows.Values.OrderBy(r => r.RowIndex))
                {
                    sheetDataElement.AppendChild(row);
                }

                worksheet.AppendChild(sheetDataElement);
                worksheetPart.Worksheet = worksheet;

                // Add merged cells if any
                if (mergedCells.Count > 0)
                {
                    var mergeCellsElement = new MergeCells();
                    foreach (var mergeCell in mergedCells)
                    {
                        mergeCellsElement.AppendChild(mergeCell);
                    }
                    worksheet.AppendChild(mergeCellsElement);
                }

                // Add sheet to workbook
                var sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = (uint)(sheets.Count() + 1),
                    Name = sheetData.Key
                };
                sheets.AppendChild(sheet);
            }
        }

        private static (uint rowIndex, string columnIndex) ParseCellAddress(string cellAddress)
        {
            var match = CellAddressRegex().Match(cellAddress);
            if (match.Success)
            {
                var column = match.Groups[1].Value;
                var row = uint.Parse(match.Groups[2].Value);
                return (row, column);
            }
            throw new ArgumentException($"Invalid cell address: {cellAddress}");
        }

        private static Stylesheet CreateStylesheet()
        {
            var stylesheet = new Stylesheet();

            // Fonts
            var fonts = new Fonts() { Count = 1 };
            fonts.AppendChild(new Font());
            stylesheet.AppendChild(fonts);

            // Fills
            var fills = new Fills() { Count = 1 };
            fills.AppendChild(new Fill());
            stylesheet.AppendChild(fills);

            // Borders
            var borders = new Borders() { Count = 1 };
            borders.AppendChild(new Border());
            stylesheet.AppendChild(borders);

            // Cell formats
            var cellFormats = new CellFormats() { Count = 1 };
            cellFormats.AppendChild(new CellFormat());
            stylesheet.AppendChild(cellFormats);

            return stylesheet;
        }

        private static uint GetOrCreateStyleIndex(Stylesheet stylesheet, CellStyleModel cellStyle)
        {
            // For simplicity, we'll create a basic style index
            // In a more complete implementation, you would check for existing styles
            // and create new ones as needed
            
            var cellFormats = stylesheet.CellFormats;
            var cellFormat = new CellFormat();

            // Apply font formatting
            if (!string.IsNullOrEmpty(cellStyle.FontName) || cellStyle.FontSize > 0 || cellStyle.Bold || cellStyle.Italic)
            {
                var font = new Font();
                
                if (!string.IsNullOrEmpty(cellStyle.FontName))
                    font.FontName = new FontName() { Val = cellStyle.FontName };
                
                if (cellStyle.FontSize > 0)
                    font.FontSize = new FontSize() { Val = cellStyle.FontSize };
                
                if (cellStyle.Bold)
                    font.Bold = new Bold();
                
                if (cellStyle.Italic)
                    font.Italic = new Italic();
                
                if (!string.IsNullOrEmpty(cellStyle.FontColor))
                    font.Color = new Color() { Rgb = cellStyle.FontColor };

                // Add font to stylesheet and reference it
                var fonts = stylesheet.Fonts;
                fonts.AppendChild(font);
                fonts.Count = (uint)fonts.Count();
                cellFormat.FontId = (uint)(fonts.Count() - 1);
            }

            // Apply fill formatting
            if (cellStyle.FillForegroundColor != null || cellStyle.FillBackgroundColor != null)
            {
                var fill = new Fill();
                var patternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid
                };

                if (!string.IsNullOrEmpty(cellStyle.FillForegroundColor))
                {
                    patternFill.ForegroundColor = new ForegroundColor()
                    {
                        Rgb = cellStyle.FillForegroundColor
                    };
                }

                if (!string.IsNullOrEmpty(cellStyle.FillBackgroundColor))
                {
                    patternFill.BackgroundColor = new BackgroundColor()
                    {
                        Rgb = cellStyle.FillBackgroundColor
                    };
                }
                fill.AppendChild(patternFill);

                // Add fill to stylesheet and reference it
                var fills = stylesheet.Fills;
                fills.AppendChild(fill);
                fills.Count = (uint)fills.Count();
                cellFormat.FillId = (uint)(fills.Count() - 1);
            }

            // Apply alignment
            if (!string.IsNullOrEmpty(cellStyle.HorizontalAlign) || !string.IsNullOrEmpty(cellStyle.VerticalAlign))
            {
                var alignment = new Alignment();
                var horizontalAlign = new HorizontalAlignmentValues(cellStyle.HorizontalAlign);
                var verticalAlign = new VerticalAlignmentValues(cellStyle.VerticalAlign);

                if ((horizontalAlign as IEnumValue).IsValid)
                {
                    
                    alignment.Horizontal = horizontalAlign;
                }
                
                if ((verticalAlign as IEnumValue).IsValid)
                {
                    alignment.Vertical = verticalAlign;
                }
                
                cellFormat.Alignment = alignment;
            }

            // Add cell format to stylesheet
            cellFormats.AppendChild(cellFormat);
            cellFormats.Count = (uint)cellFormats.Count() + 1;
            return (uint)(cellFormats.Count() - 1);
        }

        [System.Text.RegularExpressions.GeneratedRegex(@"([A-Z]+)(\d+)")]
        private static partial System.Text.RegularExpressions.Regex CellAddressRegex();
    }
}
