using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using excelConvert.ViewModels;

namespace excelConvert
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindowViewModel ViewModel { get; set; }
        
        public MainWindow()
        {
            InitializeComponent();
            ViewModel = new MainWindowViewModel();
            DataContext = ViewModel;
        }
        
        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (ViewModel != null)
            {
                ViewModel.SelectedTreeItem = e.NewValue;
            }
        }
        
        private void TreeView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // 获取被双击的元素
            var treeViewItem = FindTreeViewItem((DependencyObject)e.OriginalSource);
            if (treeViewItem != null && treeViewItem.DataContext != null)
            {
                // 检查是否是ExcelSheetItem
                if (treeViewItem.DataContext is ViewModels.ExcelSheetItem sheetItem)
                {
                    // 调用ViewModel的ExportConfig方法
                    ViewModel.ExportConfigCommand.Execute(sheetItem);
                }
            }
        }
        
        private TreeViewItem FindTreeViewItem(DependencyObject dependencyObject)
        {
            if (dependencyObject is TreeViewItem)
                return dependencyObject as TreeViewItem;
            
            DependencyObject parent = VisualTreeHelper.GetParent(dependencyObject);
            if (parent != null)
                return FindTreeViewItem(parent);
            
            return null;
        }
    }
}