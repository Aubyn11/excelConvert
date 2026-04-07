using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// 设置 EPPlus 许可证
ExcelPackage.License.SetNonCommercialOrganization("ExcelExporter");

// 计算项目根目录（exe 在 publish/ 下，向上 4 级到 POELike/）
string baseDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", ".."));
string excelDir = Path.Combine(baseDir, "common", "excel", "xls");
string cfgDir   = Path.Combine(baseDir, "common", "cfg");
string outputDir = Path.Combine(baseDir, "Assets", "Cfg");

Console.WriteLine($"项目根目录: {baseDir}");
Console.WriteLine($"Excel目录:  {excelDir}");
Console.WriteLine($"输出目录:   {outputDir}");
Console.WriteLine();

if (!Directory.Exists(excelDir))
{
    Console.WriteLine($"[错误] Excel目录不存在: {excelDir}");
    Console.WriteLine("按任意键退出...");
    Console.ReadKey();
    return;
}

if (!Directory.Exists(outputDir))
    Directory.CreateDirectory(outputDir);

int successCount = 0, failCount = 0;

foreach (var excelFile in Directory.GetFiles(excelDir, "*.xlsx"))
{
    Console.WriteLine($"处理文件: {Path.GetFileName(excelFile)}");
    using var pkg = new ExcelPackage(new FileInfo(excelFile));
    foreach (var ws in pkg.Workbook.Worksheets)
    {
        if (ws.Dimension == null) continue;

        var firstCell = ws.Cells[1, 1].Value?.ToString() ?? "";
        if (!firstCell.StartsWith("convert(") || !firstCell.EndsWith(")")) continue;

        var content = firstCell.Substring(8, firstCell.Length - 9);
        var parts = content.Split(',');
        if (parts.Length != 3) continue;

        var configFile = parts[0].Trim();
        var exportFile = parts[1].Trim();
        var dataScheme = parts[2].Trim();

        var configPath = Path.Combine(cfgDir, configFile);
        if (!File.Exists(configPath))
        {
            Console.WriteLine($"  [跳过] 配置文件不存在: {configFile}");
            failCount++;
            continue;
        }

        var fields = ParseFields(File.ReadAllText(configPath), dataScheme);
        int headerRow = 2;
        var colMap = new Dictionary<string, int>();
        for (int c = 1; c <= ws.Dimension.Columns; c++)
        {
            var colName = ws.Cells[headerRow, c].Value?.ToString();
            if (!string.IsNullOrEmpty(colName)) colMap[colName] = c;
        }

        var dataList = new List<Dictionary<string, object>>();
        for (int r = headerRow + 1; r <= ws.Dimension.Rows; r++)
        {
            var row = new Dictionary<string, object>();
            foreach (var (name, isRepeated, _) in fields)
            {
                if (!colMap.TryGetValue(name, out int colIdx))
                {
                    row[name] = isRepeated ? (object)new List<int>() : "";
                    continue;
                }
                var val = ws.Cells[r, colIdx].Value?.ToString() ?? "";
                if (isRepeated)
                {
                    var arr = new List<int>();
                    foreach (var p in val.Split(','))
                        if (int.TryParse(p.Trim(), out int iv)) arr.Add(iv);
                    row[name] = arr;
                }
                else
                {
                    row[name] = val;
                }
            }
            dataList.Add(row);
        }

        string cleanExportFile = exportFile.EndsWith(".pb", StringComparison.OrdinalIgnoreCase)
            ? exportFile : $"{exportFile}.pb";
        var outputPath = Path.Combine(outputDir, cleanExportFile);
        var exportData = new Dictionary<string, object> { [dataScheme] = dataList };
        var json = JsonSerializer.Serialize(exportData, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(outputPath, json);
        Console.WriteLine($"  [成功] {ws.Name} => {cleanExportFile}");
        successCount++;
    }
}

Console.WriteLine();
Console.WriteLine($"导出完成: 成功 {successCount} 个，失败 {failCount} 个");
if (Environment.UserInteractive && !Console.IsInputRedirected)
{
    Console.WriteLine("按任意键退出...");
    Console.ReadKey();
}

static List<(string name, bool isRepeated, string type)> ParseFields(string content, string dataScheme)
{
    var result = new List<(string, bool, string)>();
    var lines = content.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
    bool inTarget = false;
    foreach (var line in lines)
    {
        var t = line.Trim();
        if (t.EndsWith("{")) { inTarget = t.Substring(0, t.Length - 1).Trim() == dataScheme; }
        else if (t == "}") inTarget = false;
        else if (inTarget && !string.IsNullOrEmpty(t))
        {
            var fp = t.Split(',')[0].Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (fp.Length >= 3 && fp[0].ToLower() == "repeated") result.Add((fp[2], true, fp[1]));
            else if (fp.Length >= 2) result.Add((fp[1], false, fp[0]));
        }
    }
    return result;
}