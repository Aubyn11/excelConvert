using excelConvert.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace excelConvert.Services
{
    public class ExcelService
    {
        // 静态构造函数，在类加载时执行，并且只执行一次
        static ExcelService()
        {
            // 设置EPPlus的许可证为非商业用途
            try
            {
                // 使用EPPlus 8推荐的方法设置许可证
                ExcelPackage.License.SetNonCommercialOrganization("excelConvert Application");
            }
            catch (Exception ex)
            {
                // 忽略设置许可证时的错误
                System.Windows.MessageBox.Show($"设置EPPlus许可证时发生错误: {ex.Message}", "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }
        
        public bool OpenExcelFile(string filePath, out List<List<string>> data, out string errorMessage)
        {
            data = new List<List<string>>();
            errorMessage = string.Empty;
            
            try
            {
                // 检查文件是否存在
                if (!File.Exists(filePath))
                {
                    errorMessage = "文件不存在";
                    MessageBox.Show(errorMessage, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                
                // 创建文件信息对象
                var fileInfo = new FileInfo(filePath);
                
                using var package = new ExcelPackage(fileInfo);
                
                // 检查Workbook是否为null
                if (package.Workbook == null)
                {
                    errorMessage = "Excel文件格式不正确，无法读取Workbook";
                    MessageBox.Show(errorMessage, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                
                // 检查是否有工作表
                if (package.Workbook.Worksheets.Count == 0)
                {
                    errorMessage = "Excel文件中没有工作表";
                    MessageBox.Show(errorMessage, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                
                // 获取第一个工作表（使用索引0，因为EPPlus的工作表索引从0开始）
                var worksheet = package.Workbook.Worksheets[0];
                
                // 检查工作表是否有数据
                if (worksheet.Dimension == null)
                {
                    errorMessage = "工作表中没有数据";
                    MessageBox.Show(errorMessage, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;
                
                // 读取所有数据
                for (int row = 1; row <= rowCount; row++)
                {
                    List<string> rowData = new List<string>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;
                        rowData.Add(cellValue?.ToString() ?? string.Empty);
                    }
                    data.Add(rowData);
                }
                
                MessageBox.Show($"成功读取 {data.Count} 行数据", "信息", MessageBoxButton.OK, MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"打开Excel文件失败: {ex.Message}";
                MessageBox.Show(errorMessage, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }
        
        public bool OpenExcelFileWithConfig(string filePath, Config config, out Dictionary<string, List<Dictionary<string, string>>> data, out string errorMessage)
        {
            data = new Dictionary<string, List<Dictionary<string, string>>>();
            errorMessage = string.Empty;
            
            try
            {
                // 检查文件是否存在
                if (!File.Exists(filePath))
                {
                    errorMessage = "文件不存在";
                    MessageBox.Show(errorMessage, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                
                using var package = new ExcelPackage(new FileInfo(filePath));
                
                // 检查Workbook是否为null
                if (package.Workbook == null)
                {
                    errorMessage = "Excel文件格式不正确，无法读取Workbook";
                    MessageBox.Show(errorMessage, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                
                // 检查是否有工作表
                if (package.Workbook.Worksheets.Count == 0)
                {
                    errorMessage = "Excel文件中没有工作表";
                    MessageBox.Show(errorMessage, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                
                // 获取第一个工作表（使用索引0，因为EPPlus的工作表索引从0开始）
                var worksheet = package.Workbook.Worksheets[0];
                
                // 检查工作表是否有数据
                if (worksheet.Dimension == null)
                {
                    errorMessage = "工作表中没有数据";
                    MessageBox.Show(errorMessage, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                
                // 获取表头行
                var headerRow = 1;
                
                // 创建列名到列索引的映射
                Dictionary<string, int> columnMap = new Dictionary<string, int>();
                int colCount = worksheet.Dimension.Columns;
                for (int col = 1; col <= colCount; col++)
                {
                    var columnName = worksheet.Cells[headerRow, col].Value?.ToString();
                    if (!string.IsNullOrEmpty(columnName))
                    {
                        columnMap[columnName] = col;
                    }
                }
                
                // 处理每个数据类型
                if (config.Sheet1?.DataTypes != null)
                {
                    foreach (var dataType in config.Sheet1.DataTypes)
                    {
                        List<Dictionary<string, string>> typeData = new List<Dictionary<string, string>>();
                        
                        // 读取数据行
                        int rowCount = worksheet.Dimension.Rows;
                        for (int row = headerRow + 1; row <= rowCount; row++)
                        {
                            Dictionary<string, string> rowData = new Dictionary<string, string>();
                            bool hasData = false;
                            
                            // 根据配置文件中的字段映射读取数据
                            foreach (var field in dataType.Fields)
                            {
                                if (columnMap.TryGetValue(field.ExcelColumn, out int colIndex))
                                {
                                    var cellValue = worksheet.Cells[row, colIndex].Value;
                                    rowData[field.ExportName] = cellValue?.ToString() ?? string.Empty;
                                    if (!string.IsNullOrEmpty(rowData[field.ExportName]))
                                    {
                                        hasData = true;
                                    }
                                }
                                else
                                {
                                    // 如果列不存在，添加空值
                                    rowData[field.ExportName] = string.Empty;
                                }
                            }
                            
                            // 总是添加数据行到结果中，即使所有字段都是空的
                            typeData.Add(rowData);
                        }
                        
                        data[dataType.Name] = typeData;
                    }
                }
                
                MessageBox.Show($"成功读取Excel文件并处理配置", "信息", MessageBoxButton.OK, MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"根据配置文件打开Excel文件失败: {ex.Message}";
                MessageBox.Show(errorMessage, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }
        
        public string GetExcelFileName(string filePath)
        {
            return Path.GetFileName(filePath);
        }
        
        public string GetConfigFileName(string excelFileName)
        {
            // 移除Excel文件扩展名，添加.json扩展名
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelFileName);
            return $"{fileNameWithoutExt}.json";
        }
    }
}