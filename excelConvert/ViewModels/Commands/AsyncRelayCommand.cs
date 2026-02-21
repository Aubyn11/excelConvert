using System;
using System.Threading.Tasks;
using System.Windows.Input;

namespace excelConvert.ViewModels.Commands
{
    /// <summary>
    /// 异步命令的实现，用于处理异步操作的命令逻辑
    /// </summary>
    public class AsyncRelayCommand : ICommand
    {
        /// <summary>
        /// 执行异步命令的委托
        /// </summary>
        private readonly Func<object, Task> _execute;
        /// <summary>
        /// 检查命令是否可以执行的委托
        /// </summary>
        private readonly Func<object, bool> _canExecute;
        /// <summary>
        /// 命令是否正在执行
        /// </summary>
        private bool _isExecuting;
        
        /// <summary>
        /// 命令可执行状态变更事件
        /// </summary>
        public event EventHandler CanExecuteChanged;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="execute">执行异步命令的委托</param>
        /// <param name="canExecute">检查命令是否可以执行的委托，默认为null（始终可以执行）</param>
        public AsyncRelayCommand(Func<object, Task> execute, Func<object, bool> canExecute = null)
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
            return !_isExecuting && (_canExecute == null || _canExecute(parameter));
        }
        
        /// <summary>
        /// 执行异步命令
        /// </summary>
        /// <param name="parameter">命令参数</param>
        public async void Execute(object parameter)
        {
            if (!CanExecute(parameter))
                return;
            
            try
            {
                _isExecuting = true;
                RaiseCanExecuteChanged();
                await _execute(parameter);
            }
            finally
            {
                _isExecuting = false;
                RaiseCanExecuteChanged();
            }
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
    /// 泛型异步命令的实现，用于强类型的命令参数
    /// </summary>
    /// <typeparam name="T">命令参数类型</typeparam>
    public class AsyncRelayCommand<T> : ICommand
    {
        /// <summary>
        /// 执行异步命令的委托
        /// </summary>
        private readonly Func<T, Task> _execute;
        /// <summary>
        /// 检查命令是否可以执行的委托
        /// </summary>
        private readonly Func<T, bool> _canExecute;
        /// <summary>
        /// 命令是否正在执行
        /// </summary>
        private bool _isExecuting;
        
        /// <summary>
        /// 命令可执行状态变更事件
        /// </summary>
        public event EventHandler CanExecuteChanged;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="execute">执行异步命令的委托</param>
        /// <param name="canExecute">检查命令是否可以执行的委托，默认为null（始终可以执行）</param>
        public AsyncRelayCommand(Func<T, Task> execute, Func<T, bool> canExecute = null)
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
            return !_isExecuting && (_canExecute == null || _canExecute((T)parameter));
        }
        
        /// <summary>
        /// 执行异步命令
        /// </summary>
        /// <param name="parameter">命令参数</param>
        public async void Execute(object parameter)
        {
            if (!CanExecute(parameter))
                return;
            
            try
            {
                _isExecuting = true;
                RaiseCanExecuteChanged();
                await _execute((T)parameter);
            }
            finally
            {
                _isExecuting = false;
                RaiseCanExecuteChanged();
            }
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