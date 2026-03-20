using excelConvert.Models;
using excelConvert.Services;
using excelConvert.ViewModels.Commands;

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Input;

namespace excelConvert.ViewModels
{
    /// <summary>
    /// Excel文件的数据模型
    /// </summary>
    public class ExcelFileItem
    {
        /// <summary>
        /// Excel文件名称
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Excel文件路径
        /// </summary>
        public string Path { get; set; }
        /// <summary>
        /// Excel文件包含的工作表列表
        /// </summary>
        public List<ExcelSheetItem> Sheets { get; set; }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        public ExcelFileItem()
        {
            Sheets = new List<ExcelSheetItem>();
        }
    }
    
    /// <summary>
    /// Excel工作表的数据模型
    /// </summary>
    public class ExcelSheetItem
    {
        /// <summary>
        /// 工作表名称
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 工作表所属的Excel文件
        /// </summary>
        public ExcelFileItem ParentFile { get; set; }
    }
    
    /// <summary>
    /// 主窗口的ViewModel，负责处理Excel文件的加载、数据预览和导出配置等功能
    /// </summary>
    public class MainWindowViewModel : BaseViewModel
    {
        /// <summary>
        /// Excel服务，用于处理Excel文件的读取和操作
        /// </summary>
        private readonly ExcelService _excelService;
        /// <summary>
        /// 配置服务，用于加载和处理配置文件
        /// </summary>
        private readonly ConfigService _configService;
        /// <summary>
        /// Excel文件目录路径
        /// </summary>
        private readonly string _excelDirectory;
        
        /// <summary>
        /// 状态消息，用于显示在界面上
        /// </summary>
        private string _statusMessage;
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }
        
        /// <summary>
        /// 原始Excel数据
        /// </summary>
        private List<List<string>> _excelData;
        public List<List<string>> ExcelData
        {
            get => _excelData;
            set => SetProperty(ref _excelData, value);
        }
        
        /// <summary>
        /// 根据配置文件处理后的Excel数据
        /// </summary>
        private Dictionary<string, List<Dictionary<string, string>>> _configBasedExcelData;
        public Dictionary<string, List<Dictionary<string, string>>> ConfigBasedExcelData
        {
            get => _configBasedExcelData;
            set => SetProperty(ref _configBasedExcelData, value);
        }
        
        /// <summary>
        /// 数据预览内容
        /// </summary>
        private List<string> _previewData;
        public List<string> PreviewData
        {
            get => _previewData;
            set => SetProperty(ref _previewData, value);
        }
        
        /// <summary>
        /// Excel文件树结构，用于在界面上显示Excel文件和工作表
        /// </summary>
        private List<ExcelFileItem> _excelFileTree;
        public List<ExcelFileItem> ExcelFileTree
        {
            get => _excelFileTree;
            set => SetProperty(ref _excelFileTree, value);
        }
        
        /// <summary>
        /// 选中的树节点
        /// </summary>
        private object _selectedTreeItem;
        public object SelectedTreeItem
        {
            get => _selectedTreeItem;
            set 
            {
                if (SetProperty(ref _selectedTreeItem, value))
                {
                    // 选中sheet时显示数据
                    if (value is ExcelSheetItem sheetItem)
                    {
                        ShowSheetData(sheetItem.ParentFile.Path, sheetItem.Name);
                    }
                }
            }
        }
        
        /// <summary>
        /// 是否正在执行操作
        /// </summary>
        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                if (SetProperty(ref _isBusy, value))
                {
                    // 通知所有命令状态变化
                    ((RelayCommand)ExportConfigCommand).RaiseCanExecuteChanged();
                }
            }
        }
        
        /// <summary>
        /// 导出配置命令
        /// </summary>
        public ICommand ExportConfigCommand { get; }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        public MainWindowViewModel()
        {
            // 设置EPPlus的许可证为非商业用途
            OfficeOpenXml.ExcelPackage.License.SetNonCommercialOrganization("excelConvert Application");
            
            // 从依赖注入容器获取服务
            _excelService = Services.ServiceProvider.GetService<ExcelService>();
            _configService = Services.ServiceProvider.GetService<ConfigService>();
            
            // 设置Excel文件目录
            string currentDir = Directory.GetCurrentDirectory();
            
            // 方式1：从应用程序目录向上7级到项目根目录
            string path1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "..", "..", "..", "common", "excel", "xls");
            path1 = Path.GetFullPath(path1);
            
            // 方式2：从当前工作目录向上查找
            string path2 = Path.Combine(currentDir, "..", "common", "excel", "xls");
            path2 = Path.GetFullPath(path2);
            
            // 方式3：直接使用绝对路径（基于用户提供的目录结构）
            string path3 = "c:\\Wpf\\excelConvert\\common\\excel\\xls";
            
            // 选择存在的路径
            if (Directory.Exists(path1))
            {
                _excelDirectory = path1;
            }
            else if (Directory.Exists(path2))
            {
                _excelDirectory = path2;
            }
            else if (Directory.Exists(path3))
            {
                _excelDirectory = path3;
            }
            else
            {
                // 如果都不存在，使用路径3作为默认值
                _excelDirectory = path3;
            }
            
            ExportConfigCommand = new RelayCommand(ExportConfig, CanExecuteCommand);
            StatusMessage = "准备就绪";
            PreviewData = new List<string>();
            IsBusy = false;
            
            // 读取Excel文件内容进行测试
            Services.ExcelReaderTest.ReadExcelFile();
            
            // 加载Excel文件和sheet
            LoadExcelFiles();
        }
        
        /// <summary>
        /// 加载Excel文件和sheet
        /// </summary>
        private void LoadExcelFiles()
        {
            try
            {
                IsBusy = true;
                StatusMessage = "正在加载Excel文件...";
                
                List<ExcelFileItem> fileItems = new List<ExcelFileItem>();
                
                // 检查Excel目录是否存在
                if (Directory.Exists(_excelDirectory))
                {
                    // 获取所有Excel文件（包括.xlsx和.xls）
                    string[] excelFiles = Directory.GetFiles(_excelDirectory, "*.xlsx", SearchOption.AllDirectories);
                    string[] excelFilesXls = Directory.GetFiles(_excelDirectory, "*.xls", SearchOption.AllDirectories);
                    
                    // 合并文件列表
                    List<string> allExcelFiles = new List<string>(excelFiles);
                    allExcelFiles.AddRange(excelFilesXls);
                    
                    foreach (string filePath in allExcelFiles)
                    {
                        ExcelFileItem fileItem = new ExcelFileItem
                        {
                            Name = Path.GetFileName(filePath),
                            Path = filePath
                        };
                        
                        // 获取文件的sheet列表
                        List<string> sheetNames = _excelService.GetExcelSheetNames(filePath);
                        
                        foreach (string sheetName in sheetNames)
                        {
                            fileItem.Sheets.Add(new ExcelSheetItem
                            {
                                Name = sheetName,
                                ParentFile = fileItem
                            });
                        }
                        
                        fileItems.Add(fileItem);
                    }
                }
                
                ExcelFileTree = fileItems;
                StatusMessage = $"已加载 {fileItems.Count} 个Excel文件";
            }
            catch (Exception ex)
            {
                Services.ExceptionHandler.HandleException(ex, "加载Excel文件时发生错误");
                StatusMessage = "加载Excel文件时发生错误，请查看详细信息";
            }
            finally
            {
                IsBusy = false;
            }
        }
        
        /// <summary>
        /// 显示sheet数据
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="sheetName">工作表名称</param>
        private void ShowSheetData(string filePath, string sheetName)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(sheetName))
                {
                    return;
                }
                
                // 读取原始Excel数据
                if (_excelService.ReadSheetData(filePath, sheetName, out List<List<string>> data, out string errorMessage))
                {
                    ExcelData = data;
                    
                    // 尝试从Excel工作表中获取配置信息
                    string configFileName, exportFileName, dataScheme;
                    bool hasConfigInfo = _excelService.GetConfigInfoFromSheet(filePath, sheetName, out configFileName, out exportFileName, out dataScheme);
                    
                    // 如果没有配置信息，使用默认的配置文件名称
                    if (!hasConfigInfo)
                    {
                        string fileName = _excelService.GetExcelFileName(filePath);
                        configFileName = _excelService.GetConfigFileName(fileName);
                        dataScheme = string.Empty;
                    }
                    
                    if (_configService.LoadConfig(configFileName, dataScheme, out Config config))
                        {
                            // 使用配置文件读取Excel数据
                            if (_excelService.OpenExcelFileWithConfig(filePath, sheetName, config, out Dictionary<string, List<Dictionary<string, string>>> configData, out string configError))
                            {
                                ConfigBasedExcelData = configData;
                                StatusMessage = $"已使用配置文件: {configFileName} 处理数据";
                            }
                            else
                            {
                                StatusMessage = $"使用配置文件处理数据失败: {configError}";
                                ConfigBasedExcelData = null;
                            }
                        }
                    else
                    {
                        StatusMessage = $"未找到配置文件: {configFileName}，显示原始数据";
                        ConfigBasedExcelData = null;
                    }
                    
                    // 自动更新数据预览
                    DataPreview(null);
                }
                else
                {
                    StatusMessage = $"读取工作表数据失败: {errorMessage}";
                    ExcelData = null;
                    ConfigBasedExcelData = null;
                }
            }
            catch (Exception ex)
            {
                Services.ExceptionHandler.HandleException(ex, "读取工作表数据时发生错误");
                StatusMessage = "读取工作表数据时发生错误，请查看详细信息";
            }
        }
        
        /// <summary>
        /// 检查命令是否可以执行
        /// </summary>
        /// <param name="parameter">命令参数</param>
        /// <returns>如果命令可以执行则返回true，否则返回false</returns>
        private bool CanExecuteCommand(object parameter)
        {
            return !IsBusy;
        }
        
        /// <summary>
        /// 生成数据预览
        /// </summary>
        /// <param name="parameter">命令参数</param>
        private void DataPreview(object parameter = null)
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
        
        /// <summary>
        /// 导出配置文件
        /// </summary>
        /// <param name="parameter">命令参数，应为ExcelSheetItem类型</param>
        private void ExportConfig(object parameter)
        {
            try
            {
                IsBusy = true;
                
                // 处理从TreeView传递过来的ExcelSheetItem参数
                ExcelSheetItem sheetItem = null;
                if (parameter is ExcelSheetItem)
                {
                    sheetItem = parameter as ExcelSheetItem;
                }
                else if (SelectedTreeItem is ExcelSheetItem)
                {
                    sheetItem = SelectedTreeItem as ExcelSheetItem;
                }
                
                // 如果没有选中sheet，提示错误
                if (sheetItem == null)
                {
                    StatusMessage = "请先选择一个Excel工作表";
                    return;
                }
                
                // 确保sheet数据已加载
                ShowSheetData(sheetItem.ParentFile.Path, sheetItem.Name);
                
                // 检查是否有数据可导出
                if (ConfigBasedExcelData == null && ExcelData == null)
                {
                    StatusMessage = "没有数据可导出";
                    return;
                }
                
                // 尝试从Excel工作表中获取配置信息
                string configFileName, exportFileName, dataScheme;
                bool hasConfigInfo = _excelService.GetConfigInfoFromSheet(sheetItem.ParentFile.Path, sheetItem.Name, out configFileName, out exportFileName, out dataScheme);
                
                // 使用配置信息中的导出文件名，如果没有则使用默认文件名
                string defaultFileName = hasConfigInfo && !string.IsNullOrEmpty(exportFileName) ? exportFileName : $"{Path.GetFileNameWithoutExtension(sheetItem.ParentFile.Name)}_{sheetItem.Name}";
                
                // 计算 Assets/Cfg 目录路径（从exe向上7级到项目根目录，再进入Assets/Cfg）
                string cfgDirectory = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "..", "..", "..", "Assets", "Cfg"));
                
                // 如果目录不存在则创建
                if (!Directory.Exists(cfgDirectory))
                {
                    Directory.CreateDirectory(cfgDirectory);
                }
                
                // 避免重复添加.pb后缀
                string cleanFileName = defaultFileName.EndsWith(".pb", StringComparison.OrdinalIgnoreCase)
                    ? defaultFileName
                    : $"{defaultFileName}.pb";
                string filePath = Path.Combine(cfgDirectory, cleanFileName);
                
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