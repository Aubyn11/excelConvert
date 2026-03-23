using excelConvert.Services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows;

namespace excelConvert
{
    public partial class App : Application
    {
        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

        /// <summary>
        /// 自定义入口点：--export-all 模式下完全不启动 WPF，直接执行批量导出
        /// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
            if (args != null && Array.Exists(args, a => a == "--export-all"))
            {
                AllocConsole();
                // 设置 EPPlus 许可证
                try { ExcelPackage.License.SetNonCommercialOrganization("excelConvert Application"); } catch { }
                BatchExportAllStatic();
                return;
            }

            // 正常启动 WPF
            var app = new App();
            app.InitializeComponent();
            app.Run();
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // 设置EPPlus的许可证为非商业用途
            try { ExcelPackage.License.SetNonCommercialOrganization("excelConvert Application"); } catch { }

            // 初始化依赖注入容器
            Services.ServiceProvider.Initialize();

            // 初始化异常处理程序
            Services.ExceptionHandler.Initialize();

            // 更新ConfigModels.cs文件
            UpdateConfigModels();

            // 检查配置文件
            CheckConfigFiles();

            // 手动启动主窗口
            var mainWindow = new MainWindow();
            mainWindow.Show();
        }

        private void UpdateConfigModels()
        {
            try
            {
                var generator = Services.ServiceProvider.GetService<ConfigModelGenerator>();
                bool updated = generator.UpdateConfigModels();
                if (updated)
                    MessageBox.Show("ConfigModels.cs文件已更新，请重新构建项目。", "信息", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新ConfigModels.cs文件时发生错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CheckConfigFiles()
        {
            try
            {
                string configDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cfg");
                if (Directory.Exists(configDirectory))
                {
                    string[] configFiles = Directory.GetFiles(configDirectory, "*.json");
                    var configService = Services.ServiceProvider.GetService<ConfigService>();
                    foreach (string configFile in configFiles)
                    {
                        string fileName = Path.GetFileName(configFile);
                        bool success = configService.LoadConfig(fileName, string.Empty, out var config);
                        if (!success)
                            MessageBox.Show($"配置文件 {fileName} 结构不正确，请检查并修改。", "配置文件错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"检查配置文件时发生错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 批量导出所有 Excel 文件到 Assets/Cfg 目录（静态方法，不依赖 WPF）
        /// </summary>
        private static void BatchExportAllStatic()
        {
            try
            {
                string baseDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", ".."));
                string excelDir = Path.Combine(baseDir, "common", "excel", "xls");
                string cfgDir = Path.Combine(baseDir, "common", "cfg");
                string outputDir = Path.Combine(baseDir, "Assets", "Cfg");

                Console.WriteLine($"项目根目录: {baseDir}");
                Console.WriteLine($"Excel目录: {excelDir}");
                Console.WriteLine($"输出目录: {outputDir}");

                if (!Directory.Exists(excelDir)) { Console.WriteLine($"[错误] Excel目录不存在: {excelDir}"); return; }
                if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

                int successCount = 0, failCount = 0;
                foreach (var excelFile in Directory.GetFiles(excelDir, "*.xlsx"))
                {
                    Console.WriteLine($"处理: {Path.GetFileName(excelFile)}");
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
                        if (!File.Exists(configPath)) { Console.WriteLine($"  [跳过] 配置文件不存在: {configFile}"); failCount++; continue; }
                        var fields = ParseFieldsForBatch(File.ReadAllText(configPath), dataScheme);
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
                                if (!colMap.TryGetValue(name, out int colIdx)) { row[name] = isRepeated ? (object)new List<int>() : ""; continue; }
                                var val = ws.Cells[r, colIdx].Value?.ToString() ?? "";
                                if (isRepeated)
                                {
                                    var arr = new List<int>();
                                    foreach (var p in val.Split(','))
                                        if (int.TryParse(p.Trim(), out int iv)) arr.Add(iv);
                                    row[name] = arr;
                                }
                                else row[name] = val;
                            }
                            dataList.Add(row);
                        }
                        // 确保导出文件名有 .pb 后缀
                        string cleanExportFile = exportFile.EndsWith(".pb", StringComparison.OrdinalIgnoreCase) ? exportFile : $"{exportFile}.pb";
                        var outputPath = Path.Combine(outputDir, cleanExportFile);
                        var exportData = new Dictionary<string, object> { [dataScheme] = dataList };
                        var json = JsonSerializer.Serialize(exportData, new JsonSerializerOptions { WriteIndented = true });
                        File.WriteAllText(outputPath, json);
                        Console.WriteLine($"  [成功] {ws.Name} => {cleanExportFile}");
                        successCount++;
                    }
                }
                Console.WriteLine($"\n导出完成: 成功 {successCount} 个，失败 {failCount} 个");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[错误] 批量导出失败: {ex.Message}\n{ex.StackTrace}");
            }
        }

        private static List<(string name, bool isRepeated, string type)> ParseFieldsForBatch(string content, string dataScheme)
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
    }
}