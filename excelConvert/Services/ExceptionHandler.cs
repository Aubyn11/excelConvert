using System;
using System.Windows;

namespace excelConvert.Services
{
    public static class ExceptionHandler
    {
        public static void Initialize()
        {
            // 注册全局未处理异常事件
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
            
            // 注册UI线程未处理异常事件
            Application.Current.DispatcherUnhandledException += OnDispatcherUnhandledException;
        }
        
        private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception;
            if (exception != null)
            {
                LogException(exception, "全局未处理异常");
                
                // 显示错误信息
                MessageBox.Show(
                    $"发生了未处理的异常: {exception.Message}\n\n请联系开发人员获取帮助。",
                    "严重错误",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
        }
        
        private static void OnDispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            LogException(e.Exception, "UI线程未处理异常");
            
            // 显示错误信息
            MessageBox.Show(
                $"发生了应用程序错误: {e.Exception.Message}\n\n请检查操作是否正确，或联系开发人员获取帮助。",
                "应用程序错误",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            );
            
            // 标记异常为已处理，防止应用程序崩溃
            e.Handled = true;
        }
        
        public static void LogException(Exception ex, string context = "")
        {
            try
            {
                // 这里可以添加日志记录逻辑，例如写入日志文件
                Console.WriteLine($"[{DateTime.Now}] {context}: {ex.Message}\n{ex.StackTrace}");
            }
            catch
            {
                // 忽略日志记录时的错误
            }
        }
        
        public static void HandleException(Exception ex, string context = "")
        {
            LogException(ex, context);
            
            // 显示错误信息
            MessageBox.Show(
                $"操作失败: {ex.Message}\n\n{context}",
                "错误",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            );
        }
    }
}