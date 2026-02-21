using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace excelConvert.Services
{
    /// <summary>
    /// Excel读取测试类，用于调试Excel文件读取问题
    /// </summary>
    public static class ExcelReaderTest
    {
        /// <summary>
        /// 读取Excel文件内容并输出到文件
        /// </summary>
        public static void ReadExcelFile()
        {
            try
            {
                // 设置Excel文件路径
                string excelPath = "c:\\Wpf\\excelConvert\\common\\excel\\xls\\test.xlsx";
                string outputPath = "c:\\Wpf\\excelConvert\\excel_output.txt";
                
                // 清空输出文件
                File.WriteAllText(outputPath, string.Empty);
                
                WriteToFile(outputPath, $"尝试读取Excel文件: {excelPath}");
                
                if (!File.Exists(excelPath))
                {
                    WriteToFile(outputPath, "Excel文件不存在");
                    return;
                }
                
                using var package = new ExcelPackage(new FileInfo(excelPath));
                
                if (package.Workbook == null)
                {
                    WriteToFile(outputPath, "无法读取Workbook");
                    return;
                }
                
                WriteToFile(outputPath, $"Excel文件包含 {package.Workbook.Worksheets.Count} 个工作表");
                
                // 遍历所有工作表
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    WriteToFile(outputPath, $"\n工作表名称: {worksheet.Name}");
                    
                    if (worksheet.Dimension == null)
                    {
                        WriteToFile(outputPath, "工作表为空");
                        continue;
                    }
                    
                    WriteToFile(outputPath, $"工作表维度: {worksheet.Dimension.Rows} 行 x {worksheet.Dimension.Columns} 列");
                    
                    // 读取第一行第一个单元格的值
                    var firstCellValue = worksheet.Cells[1, 1].Value;
                    WriteToFile(outputPath, $"第一行第一个单元格的值: {firstCellValue}");
                    
                    // 读取前几行数据
                    int rowCount = Math.Min(worksheet.Dimension.Rows, 5); // 只读取前5行
                    int colCount = Math.Min(worksheet.Dimension.Columns, 5); // 只读取前5列
                    
                    WriteToFile(outputPath, "前几行数据:");
                    for (int row = 1; row <= rowCount; row++)
                    {
                        string rowData = "";
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value;
                            rowData += $"{cellValue?.ToString() ?? "空"}\t";
                        }
                        WriteToFile(outputPath, rowData);
                    }
                }
                
                WriteToFile(outputPath, "\nExcel文件读取测试完成");
            }
            catch (Exception ex)
            {
                string outputPath = "c:\\Wpf\\excelConvert\\excel_output.txt";
                WriteToFile(outputPath, $"读取Excel文件时发生错误: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 写入文本到文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="text">要写入的文本</param>
        private static void WriteToFile(string filePath, string text)
        {
            try
            {
                File.AppendAllText(filePath, text + Environment.NewLine);
                Console.WriteLine(text);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"写入文件时发生错误: {ex.Message}");
            }
        }
    }
}
