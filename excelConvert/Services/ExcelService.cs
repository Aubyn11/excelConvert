using excelConvert.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace excelConvert.Services
{
    /// <summary>
    /// Excel服务，用于处理Excel文件的读取和操作
    /// </summary>
    public class ExcelService
    {
        /// <summary>
        /// 静态构造函数，在类加载时执行，并且只执行一次
        /// 用于设置EPPlus的许可证
        /// </summary>
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
                Console.WriteLine($"设置EPPlus许可证时发生错误: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 打开Excel文件并读取数据
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="data">读取的数据</param>
        /// <param name="errorMessage">错误信息</param>
        /// <returns>如果读取成功则返回true，否则返回false</returns>
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
                    return false;
                }
                
                // 创建文件信息对象
                var fileInfo = new FileInfo(filePath);
                
                using var package = new ExcelPackage(fileInfo);
                
                // 检查Workbook是否为null
                if (package.Workbook == null)
                {
                    errorMessage = "Excel文件格式不正确，无法读取Workbook";
                    return false;
                }
                
                // 检查是否有工作表
                if (package.Workbook.Worksheets.Count == 0)
                {
                    errorMessage = "Excel文件中没有工作表";
                    return false;
                }
                
                // 获取第一个工作表（使用索引0，因为EPPlus的工作表索引从0开始）
                var worksheet = package.Workbook.Worksheets[0];
                
                // 检查工作表是否有数据
                if (worksheet.Dimension == null)
                {
                    errorMessage = "工作表中没有数据";
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
                
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"打开Excel文件失败: {ex.Message}";
                return false;
            }
        }
        
        /// <summary>
        /// 解析表格第一行第一个单元格中的配置信息
        /// </summary>
        /// <param name="cellValue">单元格值</param>
        /// <param name="configFileName">配置文件名称</param>
        /// <param name="exportFileName">导出配置文件时的名称</param>
        /// <param name="dataScheme">配置文件中对应数据的相关读取方案</param>
        /// <returns>如果解析成功则返回true，否则返回false</returns>
        public bool ParseConfigInfo(string cellValue, out string configFileName, out string exportFileName, out string dataScheme)
        {
            configFileName = string.Empty;
            exportFileName = string.Empty;
            dataScheme = string.Empty;
            
            if (string.IsNullOrEmpty(cellValue))
                return false;
            
            // 检查是否以convert(开头并以)结尾
            if (!cellValue.StartsWith("convert(") || !cellValue.EndsWith(")"))
                return false;
            
            // 提取括号内的内容
            string content = cellValue.Substring(8, cellValue.Length - 9);
            
            // 分割内容为三个部分
            string[] parts = content.Split(',');
            if (parts.Length != 3)
                return false;
            
            // 去除空格并赋值
            configFileName = parts[0].Trim();
            exportFileName = parts[1].Trim();
            dataScheme = parts[2].Trim();
            
            return true;
        }
        
        /// <summary>
        /// 根据配置文件打开Excel文件并读取数据
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="config">配置文件</param>
        /// <param name="data">读取的数据</param>
        /// <param name="errorMessage">错误信息</param>
        /// <returns>如果读取成功则返回true，否则返回false</returns>
        public bool OpenExcelFileWithConfig(string filePath, string sheetName, Config config, out Dictionary<string, List<Dictionary<string, string>>> data, out string errorMessage)
        {
            data = new Dictionary<string, List<Dictionary<string, string>>>();
            errorMessage = string.Empty;
            
            try
            {
                // 检查文件是否存在
                if (!File.Exists(filePath))
                {
                    errorMessage = "文件不存在";
                    return false;
                }
                
                using var package = new ExcelPackage(new FileInfo(filePath));
                
                // 检查Workbook是否为null
                if (package.Workbook == null)
                {
                    errorMessage = "Excel文件格式不正确，无法读取Workbook";
                    return false;
                }
                
                // 检查是否有工作表
                if (package.Workbook.Worksheets.Count == 0)
                {
                    errorMessage = "Excel文件中没有工作表";
                    return false;
                }
                
                // 根据工作表名称获取对应的工作表
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    errorMessage = $"工作表 '{sheetName}' 不存在";
                    return false;
                }
                
                // 检查工作表是否有数据
                if (worksheet.Dimension == null)
                {
                    errorMessage = "工作表中没有数据";
                    return false;
                }
                
                // 获取表头行
                var headerRow = 1;
                
                // 检查表格第一行第一个单元格是否包含配置信息
                string configInfoCellValue = worksheet.Cells[1, 1].Value?.ToString();
                string configFileName, exportFileName, dataScheme;
                bool hasConfigInfo = ParseConfigInfo(configInfoCellValue, out configFileName, out exportFileName, out dataScheme);
                
                // 如果有配置信息，跳过第一行
                if (hasConfigInfo)
                {
                    headerRow = 2;
                }
                
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
                                    string cellValueStr = cellValue?.ToString() ?? string.Empty;
                                    rowData[field.ExportName] = cellValueStr;
                                    if (!string.IsNullOrEmpty(cellValueStr))
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
                
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"根据配置文件打开Excel文件失败: {ex.Message}";
                return false;
            }
        }
        
        /// <summary>
        /// 获取Excel文件名称
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <returns>Excel文件名称</returns>
        public string GetExcelFileName(string filePath)
        {
            return Path.GetFileName(filePath);
        }
        
        /// <summary>
        /// 获取配置文件名称
        /// </summary>
        /// <param name="excelFileName">Excel文件名称</param>
        /// <returns>配置文件名称</returns>
        public string GetConfigFileName(string excelFileName)
        {
            // 移除Excel文件扩展名，添加.txt扩展名
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelFileName);
            return $"{fileNameWithoutExt}.txt";
        }
        
        /// <summary>
        /// 获取Excel文件中的工作表名称列表
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <returns>工作表名称列表</returns>
        public List<string> GetExcelSheetNames(string filePath)
        {
            List<string> sheetNames = new List<string>();
            
            try
            {
                using var package = new ExcelPackage(new FileInfo(filePath));
                
                if (package.Workbook != null)
                {
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        sheetNames.Add(worksheet.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取Excel表格列表失败: {ex.Message}");
            }
            
            return sheetNames;
        }
        
        /// <summary>
        /// 读取指定工作表的数据
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="data">读取的数据</param>
        /// <param name="errorMessage">错误信息</param>
        /// <returns>如果读取成功则返回true，否则返回false</returns>
        public bool ReadSheetData(string filePath, string sheetName, out List<List<string>> data, out string errorMessage)
        {
            data = new List<List<string>>();
            errorMessage = string.Empty;
            
            try
            {
                using var package = new ExcelPackage(new FileInfo(filePath));
                
                if (package.Workbook == null)
                {
                    errorMessage = "Excel文件格式不正确，无法读取Workbook";
                    return false;
                }
                
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    errorMessage = $"工作表 '{sheetName}' 不存在";
                    return false;
                }
                
                if (worksheet.Dimension == null)
                {
                    errorMessage = "工作表中没有数据";
                    return false;
                }
                
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;
                
                // 检查第一行第一个单元格是否包含配置信息
                int startRow = 1;
                string firstCellValue = worksheet.Cells[1, 1].Value?.ToString();
                string configFileName, exportFileName, dataScheme;
                if (ParseConfigInfo(firstCellValue, out configFileName, out exportFileName, out dataScheme))
                {
                    // 如果有配置信息，从第二行开始读取
                    startRow = 2;
                }
                
                for (int row = startRow; row <= rowCount; row++)
                {
                    List<string> rowData = new List<string>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;
                        rowData.Add(cellValue?.ToString() ?? string.Empty);
                    }
                    data.Add(rowData);
                }
                
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = $"读取工作表数据失败: {ex.Message}";
                return false;
            }
        }
        
        /// <summary>
        /// 从Excel工作表中获取配置信息
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="configFileName">配置文件名称</param>
        /// <param name="exportFileName">导出配置文件时的名称</param>
        /// <param name="dataScheme">配置文件中对应数据的相关读取方案</param>
        /// <returns>如果获取成功则返回true，否则返回false</returns>
        public bool GetConfigInfoFromSheet(string filePath, string sheetName, out string configFileName, out string exportFileName, out string dataScheme)
        {
            configFileName = string.Empty;
            exportFileName = string.Empty;
            dataScheme = string.Empty;
            
            try
            {
                using var package = new ExcelPackage(new FileInfo(filePath));
                
                if (package.Workbook == null)
                {
                    return false;
                }
                
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    return false;
                }
                
                if (worksheet.Dimension == null)
                {
                    return false;
                }
                
                // 检查第一行第一个单元格是否包含配置信息
                string firstCellValue = worksheet.Cells[1, 1].Value?.ToString();
                return ParseConfigInfo(firstCellValue, out configFileName, out exportFileName, out dataScheme);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取配置信息失败: {ex.Message}");
                return false;
            }
        }
    }
}