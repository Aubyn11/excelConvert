using excelConvert.Models;
using excelConvert.Services;
using excelConvert.ViewModels.Commands;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Input;

namespace excelConvert.ViewModels
{
    public class MainWindowViewModel : BaseViewModel
    {
        private readonly ExcelService _excelService;
        private readonly ConfigService _configService;
        
        private string _statusMessage;
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }
        
        private List<List<string>> _excelData;
        public List<List<string>> ExcelData
        {
            get => _excelData;
            set => SetProperty(ref _excelData, value);
        }
        
        private Dictionary<string, List<Dictionary<string, string>>> _configBasedExcelData;
        public Dictionary<string, List<Dictionary<string, string>>> ConfigBasedExcelData
        {
            get => _configBasedExcelData;
            set => SetProperty(ref _configBasedExcelData, value);
        }
        
        private List<string> _previewData;
        public List<string> PreviewData
        {
            get => _previewData;
            set => SetProperty(ref _previewData, value);
        }
        
        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                if (SetProperty(ref _isBusy, value))
                {
                    // 通知所有命令状态变化
                    ((RelayCommand)OpenExcelFileCommand).RaiseCanExecuteChanged();
                    ((RelayCommand)DataPreviewCommand).RaiseCanExecuteChanged();
                    ((RelayCommand)ExportConfigCommand).RaiseCanExecuteChanged();
                }
            }
        }
        
        public ICommand OpenExcelFileCommand { get; }
        public ICommand DataPreviewCommand { get; }
        public ICommand ExportConfigCommand { get; }
        
        public MainWindowViewModel()
        {
            // 设置EPPlus的许可证为非商业用途
            OfficeOpenXml.ExcelPackage.License.SetNonCommercialOrganization("excelConvert Application");
            
            // 从依赖注入容器获取服务
            _excelService = Services.ServiceProvider.GetService<ExcelService>();
            _configService = Services.ServiceProvider.GetService<ConfigService>();
            
            OpenExcelFileCommand = new RelayCommand(OpenExcelFile, CanExecuteCommand);
            DataPreviewCommand = new RelayCommand(DataPreview, CanExecuteCommand);
            ExportConfigCommand = new RelayCommand(ExportConfig, CanExecuteCommand);
            StatusMessage = "准备就绪";
            PreviewData = new List<string>();
            IsBusy = false;
        }
        
        private bool CanExecuteCommand(object parameter)
        {
            return !IsBusy;
        }
        
        private void OpenExcelFile(object parameter)
        {
            try
            {
                IsBusy = true;
                
                // 清空scrollViewer中的信息
                PreviewData = new List<string>();
                
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls",
                    Title = "选择Excel文件"
                };
                
                if (openFileDialog.ShowDialog() == true)
                {
                    string filePath = openFileDialog.FileName;
                    string fileName = _excelService.GetExcelFileName(filePath);
                    string configFileName = _excelService.GetConfigFileName(fileName);
                    
                    // 尝试读取对应的配置文件
                    if (_configService.LoadConfig(configFileName, out Config config))
                    {
                        // 使用配置文件读取Excel数据
                        if (_excelService.OpenExcelFileWithConfig(filePath, config, out Dictionary<string, List<Dictionary<string, string>>> configData, out string errorMessage))
                        {
                            ConfigBasedExcelData = configData;
                            StatusMessage = $"已成功打开文件: {fileName}，并使用配置文件: {configFileName} 读取数据";
                            // 自动进行数据预览
                            DataPreview(null);
                        }
                        else
                        {
                            StatusMessage = $"打开Excel文件失败: {errorMessage}";
                        }
                    }
                    else
                    {
                        // 没有找到配置文件，使用默认方式读取
                        if (_excelService.OpenExcelFile(filePath, out List<List<string>> data, out string errorMessage))
                        {
                            ExcelData = data;
                            StatusMessage = $"已成功打开文件: {fileName}，共 {data.Count} 行数据 (未找到配置文件，使用默认方式读取)";
                            // 自动进行数据预览
                            DataPreview(null);
                        }
                        else
                        {
                            StatusMessage = $"打开Excel文件失败: {errorMessage}";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Services.ExceptionHandler.HandleException(ex, "打开Excel文件时发生错误");
                StatusMessage = "打开Excel文件时发生错误，请查看详细信息";
            }
            finally
            {
                IsBusy = false;
            }
        }
        
        private void DataPreview(object parameter)
        {
            try
            {
                IsBusy = true;
                
                List<string> previewList = new List<string>();
                
                // 检查是否有配置文件读取的数据
                if (ConfigBasedExcelData != null && ConfigBasedExcelData.Count > 0)
                {
                    foreach (var kvp in ConfigBasedExcelData)
                    {
                        string dataType = kvp.Key;
                        var dataList = kvp.Value;
                        
                        // 为每个数据类型添加标题
                        previewList.Add($"=== {dataType} ===");
                        
                        // 为每条数据添加预览
                        foreach (var dataItem in dataList)
                        {
                            string itemPreview = string.Empty;
                            foreach (var field in dataItem)
                            {
                                itemPreview += $"{field.Key}: {field.Value}, ";
                            }
                            // 移除末尾的逗号和空格
                            if (itemPreview.Length > 2)
                            {
                                itemPreview = itemPreview.Substring(0, itemPreview.Length - 2);
                            }
                            previewList.Add(itemPreview);
                        }
                    }
                }
                // 检查是否有默认方式读取的数据
                else if (ExcelData != null && ExcelData.Count > 0)
                {
                    previewList.Add("=== 数据预览 ===");
                    foreach (var row in ExcelData)
                    {
                        string rowPreview = string.Join(", ", row);
                        previewList.Add(rowPreview);
                    }
                }
                else
                {
                    previewList.Add("没有数据可预览");
                }
                
                PreviewData = previewList;
                StatusMessage = $"已生成数据预览，共 {previewList.Count} 条";
            }
            catch (Exception ex)
            {
                Services.ExceptionHandler.HandleException(ex, "生成数据预览时发生错误");
                StatusMessage = "生成数据预览时发生错误，请查看详细信息";
            }
            finally
            {
                IsBusy = false;
            }
        }
        
        private void ExportConfig(object parameter)
        {
            try
            {
                IsBusy = true;
                
                // 检查是否有数据可导出
                if (ConfigBasedExcelData == null && ExcelData == null)
                {
                    StatusMessage = "没有数据可导出";
                    return;
                }
                
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "PB文件 (*.pb)|*.pb",
                    Title = "导出配置文件",
                    DefaultExt = "pb",
                    FileName = "config"
                };
                
                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    
                    try
                    {
                        // 根据数据类型选择导出内容
                        object exportData = null;
                        if (ConfigBasedExcelData != null && ConfigBasedExcelData.Count > 0)
                        {
                            exportData = ConfigBasedExcelData;
                        }
                        else if (ExcelData != null && ExcelData.Count > 0)
                        {
                            exportData = ExcelData;
                        }
                        
                        // 使用策略模式处理导出
                        var exportStrategy = Services.ExportStrategyFactory.CreateStrategy("pb");
                        exportStrategy.Export(exportData, filePath);
                        
                        StatusMessage = $"配置文件已成功导出到: {filePath}";
                    }
                    catch (Exception ex)
                    {
                        Services.ExceptionHandler.HandleException(ex, "导出配置文件时发生错误");
                        StatusMessage = "导出配置文件失败，请查看详细信息";
                    }
                }
            }
            catch (Exception ex)
            {
                Services.ExceptionHandler.HandleException(ex, "导出配置文件时发生错误");
                StatusMessage = "导出配置文件失败，请查看详细信息";
            }
            finally
            {
                IsBusy = false;
            }
        }
    }
}