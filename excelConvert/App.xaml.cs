using excelConvert.Services;
using OfficeOpenXml;
using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Windows;





namespace excelConvert
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        // 静态构造函数，在类加载时执行
        static App()
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
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
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
            
            // 初始化依赖注入容器
            Services.ServiceProvider.Initialize();
            
            // 初始化异常处理程序
            Services.ExceptionHandler.Initialize();
            
            // 更新ConfigModels.cs文件
            UpdateConfigModels();
            
            // 检查配置文件
            CheckConfigFiles();
        }
        
        private void UpdateConfigModels()
        {
            try
            {
                // 从依赖注入容器获取配置模型生成器
                var generator = Services.ServiceProvider.GetService<ConfigModelGenerator>();
                
                // 更新ConfigModels.cs文件
                bool updated = generator.UpdateConfigModels();
                
                if (updated)
                {
                    // 显示信息
                    MessageBox.Show("ConfigModels.cs文件已更新，请重新构建项目。", "信息", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                // 显示错误信息
                MessageBox.Show($"更新ConfigModels.cs文件时发生错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void CheckConfigFiles()
        {
            try
            {
                // 获取配置文件目录
                string configDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cfg");
                
                // 检查配置文件目录是否存在
                if (Directory.Exists(configDirectory))
                {
                    // 获取所有.json文件
                    string[] configFiles = Directory.GetFiles(configDirectory, "*.json");
                    
                    // 从依赖注入容器获取配置服务
                    var configService = Services.ServiceProvider.GetService<ConfigService>();
                    
                    // 检查每个配置文件
                    foreach (string configFile in configFiles)
                    {
                        string fileName = Path.GetFileName(configFile);
                        bool success = configService.LoadConfig(fileName, out var config);
                        
                        if (!success)
                        {
                            // 显示警告信息
                            MessageBox.Show($"配置文件 {fileName} 结构不正确，请检查并修改。", "配置文件错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 显示错误信息
                MessageBox.Show($"检查配置文件时发生错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

}
