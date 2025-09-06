using DER.Utility.Converters;

var filePath = "C:\\Users\\DELL\\Downloads\\Free_Status_Report_Template_ProjectManager_WLNK-1.xlsx";
using var fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
var json = ExcelConverter.ExcelToJson(fileStream);
var directory = Directory.CreateDirectory("JsonTemplates");
await File.WriteAllTextAsync($"JsonTemplates/json_template_{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}.json", json);

Console.WriteLine("Complete");
