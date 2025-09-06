# DynamicExcelReport

A .NET 8.0 console application that converts Excel files to JSON format while preserving cell formatting, styles, and merged cells.

## Features

- **Excel to JSON Conversion**: Converts Excel (.xlsx) files to structured JSON format
- **Style Preservation**: Maintains cell formatting including fonts, colors, alignment, and fills
- **Merged Cell Support**: Handles merged cells and their ranges
- **Multiple Sheet Support**: Processes all worksheets in the Excel file
- **Theme Color Support**: Converts theme colors to RGB values

## Project Structure

```
DynamicExcelReport/
├── DER.Console/           # Console application entry point
│   ├── Program.cs         # Main application logic
│   └── DER.Console.csproj # Console project file
├── DER.Utility/           # Core utility library
│   ├── Converters/
│   │   └── ExcelConverter.cs  # Excel to JSON conversion logic
│   ├── Models/
│   │   └── CellModel.cs       # Data models for Excel cells
│   └── DER.Utility.csproj     # Utility project file
└── DynamicExcelReport.sln     # Visual Studio solution file
```

## Dependencies

- **.NET 8.0**: Target framework
- **DocumentFormat.OpenXml**: For reading Excel files
- **Newtonsoft.Json**: For JSON serialization

## Usage

1. **Build the project**:
   ```bash
   dotnet build
   ```

2. **Run the application**:
   ```bash
   dotnet run --project DER.Console
   ```

3. **Update the file path** in `Program.cs` to point to your Excel file:
   ```csharp
   var filePath = "path/to/your/excel/file.xlsx";
   ```

4. **Output**: The application will generate a JSON file in the `JsonTemplates/` directory with a timestamp.

## JSON Output Format

The generated JSON follows this structure:

```json
{
  "SheetName": {
    "A1": {
      "value": "Cell Value",
      "mergedRange": "A1:B2",
      "style": {
        "fontName": "Calibri",
        "fontSize": 11,
        "bold": false,
        "italic": false,
        "fontColor": "FF000000",
        "fillForegroundColor": "FFFFFFFF",
        "fillBackgroundColor": "FFFFFFFF",
        "horizontalAlign": "Left",
        "verticalAlign": "Bottom"
      }
    }
  }
}
```

## Development

### Building from Source

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/DynamicExcelReport.git
   cd DynamicExcelReport
   ```

2. Restore dependencies:
   ```bash
   dotnet restore
   ```

3. Build the solution:
   ```bash
   dotnet build
   ```

### Running Tests

Currently, the project doesn't include unit tests. Consider adding them for better code coverage.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Future Enhancements

- [ ] Add unit tests
- [ ] Support for .xls files (older Excel format)
- [ ] Command-line arguments for input/output paths
- [ ] JSON to Excel conversion
- [ ] Batch processing of multiple files
- [ ] Configuration file support
