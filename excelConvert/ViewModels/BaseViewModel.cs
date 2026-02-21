using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace excelConvert.ViewModels
{
    /// <summary>
    /// 所有ViewModel的基类，实现了INotifyPropertyChanged接口
    /// 提供了属性变更通知的基本功能
    /// </summary>
    public class BaseViewModel : INotifyPropertyChanged
    {
        /// <summary>
        /// 属性变更事件
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;
        
        /// <summary>
        /// 触发属性变更事件
        /// </summary>
        /// <param name="propertyName">属性名，默认为调用方法的属性名</param>
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        
        /// <summary>
        /// 设置属性值并触发属性变更事件
        /// </summary>
        /// <typeparam name="T">属性类型</typeparam>
        /// <param name="field">字段引用</param>
        /// <param name="value">新值</param>
        /// <param name="propertyName">属性名，默认为调用方法的属性名</param>
        /// <returns>如果值发生变化则返回true，否则返回false</returns>
        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (Equals(field, value)) return false;
            
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }
}