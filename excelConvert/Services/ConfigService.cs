using excelConvert.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows;

namespace excelConvert.Services
{
    public class ConfigService
    {
        private readonly string _configDirectory;
        
        public ConfigService()
        {
            // 设置配置文件目录为common/cfg（位于项目根目录）
            _configDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "common", "cfg");
        }
        
        public bool LoadConfig(string configFileName, string dataScheme, out Config config)
        {
            config = null;
            
            try
            {
                // 构建配置文件路径
                string configPath = Path.Combine(_configDirectory, configFileName);
                
                // 检查配置文件是否存在
                if (!File.Exists(configPath))
                {
                    MessageBox.Show($"配置文件不存在: {configPath}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                
                // 读取配置文件内容
                string content = File.ReadAllText(configPath);
                
                // 根据文件扩展名决定解析方式
                string extension = Path.GetExtension(configFileName).ToLower();
                if (extension == ".json")
                {
                    // 解析JSON格式的配置文件
                    var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                    config = JsonSerializer.Deserialize<Config>(content, options);
                }
                else if (extension == ".txt")
                {
                    // 解析TXT格式的配置文件
                    // 这里需要根据txt文件的实际格式进行解析
                    // 暂时创建一个默认的Config对象
                    config = new Config
                    {
                        Sheet1 = new SheetConfig
                        {
                            DataTypes = new List<DataTypeConfig>()
                        }
                    };
                    
                    // 简单解析txt文件，提取数据类型和字段
                    string[] lines = content.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    List<FieldConfig> currentFields = null;
                    string currentDataType = null;
                    bool isTargetData = string.IsNullOrEmpty(dataScheme); // 如果dataScheme为空，则解析所有数据
                    
                    foreach (string line in lines)
                    {
                        string trimmedLine = line.Trim();
                        if (trimmedLine.EndsWith("{"))
                        {
                            // 开始一个新的数据类型
                            currentDataType = trimmedLine.Substring(0, trimmedLine.Length - 1).Trim();
                            // 检查是否是目标数据类型
                            if (string.IsNullOrEmpty(dataScheme) || currentDataType == dataScheme)
                            {
                                isTargetData = true;
                                currentFields = new List<FieldConfig>();
                            }
                            else
                            {
                                isTargetData = false;
                                currentFields = null;
                            }
                        }
                        else if (trimmedLine == "}")
                        {
                            // 结束一个数据类型
                            if (isTargetData && !string.IsNullOrEmpty(currentDataType) && currentFields != null)
                            {
                                config.Sheet1.DataTypes.Add(new DataTypeConfig
                                {
                                    Name = currentDataType,
                                    Description = currentDataType,
                                    Fields = currentFields,
                                    ExportFormats = new List<string> { "pb" }
                                });
                            }
                            currentDataType = null;
                            currentFields = null;
                            isTargetData = string.IsNullOrEmpty(dataScheme); // 重置为默认状态
                        }
                        else if (isTargetData && !string.IsNullOrEmpty(trimmedLine) && currentFields != null)
                        {
                            // 解析字段定义
                            string[] parts = trimmedLine.Split(',');
                            if (parts.Length > 0)
                            {
                                string fieldDefinition = parts[0].Trim();
                                string[] fieldParts = fieldDefinition.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                if (fieldParts.Length >= 2)
                                {
                                    // 提取字段类型和名称
                                    string fieldType = fieldParts[0];
                                    string fieldName = fieldParts[1];
                                    
                                    // 检查是否是关键字段
                                    bool isKey = fieldDefinition.Contains("(key)");
                                    
                                    currentFields.Add(new FieldConfig
                                    {
                                        ExportName = fieldName,
                                        ExcelColumn = fieldName,
                                        Type = fieldType,
                                        Required = isKey
                                    });
                                }
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show($"不支持的配置文件格式: {extension}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取配置文件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }
        
        public List<string> GetConfigFiles()
        {
            List<string> configFiles = new List<string>();
            
            try
            {
                // 检查配置文件目录是否存在
                if (Directory.Exists(_configDirectory))
                {
                    // 获取所有.json文件
                    string[] files = Directory.GetFiles(_configDirectory, "*.json");
                    foreach (string file in files)
                    {
                        configFiles.Add(Path.GetFileName(file));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取配置文件列表失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
            return configFiles;
        }
    }
}