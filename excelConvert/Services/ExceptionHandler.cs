using System;
using System.Windows;

namespace excelConvert.Services
{
    /// <summary>
    /// 异常处理器，用于全局异常的捕获和处理
    /// </summary>
    public static class ExceptionHandler
    {
        /// <summary>
        /// 初始化异常处理器，注册全局异常事件
        /// </summary>
        public static void Initialize()
        {
            // 注册全局未处理异常事件
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
            
            // 注册UI线程未处理异常事件
            Application.Current.DispatcherUnhandledException += OnDispatcherUnhandledException;
        }
        
        /// <summary>
        /// 处理全局未处理异常
        /// </summary>
        /// <param name="sender">事件发送者</param>
        /// <param name="e">异常事件参数</param>
        private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception;
            if (exception != null)
            {
                LogException(exception, "全局未处理异常");
            }
        }
        
        /// <summary>
        /// 处理UI线程未处理异常
        /// </summary>
        /// <param name="sender">事件发送者</param>
        /// <param name="e">异常事件参数</param>
        private static void OnDispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            LogException(e.Exception, "UI线程未处理异常");
            
            // 标记异常为已处理，防止应用程序崩溃
            e.Handled = true;
        }
        
        /// <summary>
        /// 记录异常信息
        /// </summary>
        /// <param name="ex">异常对象</param>
        /// <param name="context">异常上下文信息</param>
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
        
        /// <summary>
        /// 处理异常
        /// </summary>
        /// <param name="ex">异常对象</param>
        /// <param name="context">异常上下文信息</param>
        public static void HandleException(Exception ex, string context = "")
        {
            LogException(ex, context);
        }
    }
}