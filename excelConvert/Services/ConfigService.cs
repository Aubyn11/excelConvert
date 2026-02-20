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
            // 设置配置文件目录
            _configDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cfg");
        }
        
        public bool LoadConfig(string configFileName, out Config config)
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
                
                // 读取并解析配置文件
                string jsonContent = File.ReadAllText(configPath);
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                config = JsonSerializer.Deserialize<Config>(jsonContent, options);
                
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