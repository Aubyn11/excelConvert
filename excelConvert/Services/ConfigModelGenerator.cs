using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;

namespace excelConvert.Services
{
    public class ConfigModelGenerator
    {
        private readonly string _configDirectory;
        private readonly string _configModelsPath;
        
        public ConfigModelGenerator()
        {
            // 获取配置文件目录
            _configDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cfg");
            
            // 获取ConfigModels.cs文件路径
            _configModelsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "Models", "ConfigModels.cs");
        }
        
        public bool UpdateConfigModels()
        {
            try
            {
                // 检查配置文件目录是否存在
                if (!Directory.Exists(_configDirectory))
                {
                    return false;
                }
                
                // 检查ConfigModels.cs文件是否存在
                if (!File.Exists(_configModelsPath))
                {
                    return false;
                }
                
                // 生成新的ConfigModels.cs内容
                string newContent = GenerateConfigModelsContent();
                
                // 读取当前的ConfigModels.cs内容
                string currentContent = File.ReadAllText(_configModelsPath);
                
                // 检查内容是否不同
                if (newContent != currentContent)
                {
                    // 写入新内容
                    File.WriteAllText(_configModelsPath, newContent);
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新ConfigModels.cs文件时发生错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }
        
        private string GenerateConfigModelsContent()
        {
            StringBuilder sb = new StringBuilder();
            
            // 添加命名空间和using语句
            sb.AppendLine("using System.Collections.Generic;");
            sb.AppendLine();
            sb.AppendLine("namespace excelConvert.Models");
            sb.AppendLine("{");
            
            // 添加Config类
            sb.AppendLine("    public class Config");
            sb.AppendLine("    {");
            sb.AppendLine("        public SheetConfig Sheet1 { get; set; }");
            sb.AppendLine("    }");
            sb.AppendLine();
            
            // 添加SheetConfig类
            sb.AppendLine("    public class SheetConfig");
            sb.AppendLine("    {");
            sb.AppendLine("        public List<DataTypeConfig> DataTypes { get; set; }");
            sb.AppendLine("    }");
            sb.AppendLine();
            
            // 添加DataTypeConfig类
            sb.AppendLine("    public class DataTypeConfig");
            sb.AppendLine("    {");
            sb.AppendLine("        public string Name { get; set; }");
            sb.AppendLine("        public string Description { get; set; }");
            sb.AppendLine("        public List<FieldConfig> Fields { get; set; }");
            sb.AppendLine("        public List<string> ExportFormats { get; set; }");
            sb.AppendLine("    }");
            sb.AppendLine();
            
            // 添加FieldConfig类
            sb.AppendLine("    public class FieldConfig");
            sb.AppendLine("    {");
            sb.AppendLine("        public string ExportName { get; set; }");
            sb.AppendLine("        public string ExcelColumn { get; set; }");
            sb.AppendLine("        public bool Required { get; set; }");
            sb.AppendLine("        public string Type { get; set; }");
            sb.AppendLine("    }");
            sb.AppendLine("}");
            
            return sb.ToString();
        }
    }
}