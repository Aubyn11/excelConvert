using System;
using System.Windows.Input;

namespace excelConvert.ViewModels.Commands
{
    /// <summary>
    /// 命令的基本实现，用于将命令逻辑与UI分离
    /// </summary>
    public class RelayCommand : ICommand
    {
        /// <summary>
        /// 执行命令的委托
        /// </summary>
        private readonly Action<object> _execute;
        /// <summary>
        /// 检查命令是否可以执行的委托
        /// </summary>
        private readonly Func<object, bool> _canExecute;
        
        /// <summary>
        /// 命令可执行状态变更事件
        /// </summary>
        public event EventHandler CanExecuteChanged;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="execute">执行命令的委托</param>
        /// <param name="canExecute">检查命令是否可以执行的委托，默认为null（始终可以执行）</param>
        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }
        
        /// <summary>
        /// 检查命令是否可以执行
        /// </summary>
        /// <param name="parameter">命令参数</param>
        /// <returns>如果命令可以执行则返回true，否则返回false</returns>
        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute(parameter);
        }
        
        /// <summary>
        /// 执行命令
        /// </summary>
        /// <param name="parameter">命令参数</param>
        public void Execute(object parameter)
        {
            _execute(parameter);
        }
        
        /// <summary>
        /// 触发命令可执行状态变更事件
        /// </summary>
        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }
    
    /// <summary>
    /// 泛型命令的实现，用于强类型的命令参数
    /// </summary>
    /// <typeparam name="T">命令参数类型</typeparam>
    public class RelayCommand<T> : ICommand
    {
        /// <summary>
        /// 执行命令的委托
        /// </summary>
        private readonly Action<T> _execute;
        /// <summary>
        /// 检查命令是否可以执行的委托
        /// </summary>
        private readonly Func<T, bool> _canExecute;
        
        /// <summary>
        /// 命令可执行状态变更事件
        /// </summary>
        public event EventHandler CanExecuteChanged;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="execute">执行命令的委托</param>
        /// <param name="canExecute">检查命令是否可以执行的委托，默认为null（始终可以执行）</param>
        public RelayCommand(Action<T> execute, Func<T, bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }
        
        /// <summary>
        /// 检查命令是否可以执行
        /// </summary>
        /// <param name="parameter">命令参数</param>
        /// <returns>如果命令可以执行则返回true，否则返回false</returns>
        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute((T)parameter);
        }
        
        /// <summary>
        /// 执行命令
        /// </summary>
        /// <param name="parameter">命令参数</param>
        public void Execute(object parameter)
        {
            _execute((T)parameter);
        }
        
        /// <summary>
        /// 触发命令可执行状态变更事件
        /// </summary>
        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}