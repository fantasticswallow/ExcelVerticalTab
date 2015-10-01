using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace VerticalTabControlLib
{
    /// <summary>
    /// MyCanvas.xaml の相互作用ロジック
    /// </summary>
    public partial class TabUserControl : UserControl
    {
        public TabUserControl()
        {
            InitializeComponent();
        }
    }

    public interface ITabUserControlViewModel<T> 
    {
        ObservableCollection<T> Items { get; set; }
        T SelectedItem { get; set; }
         
    }

    internal class TabUserControlViewModelMock : ITabUserControlViewModel<object>
    {
        public ObservableCollection<object> Items { get; set; } = new ObservableCollection<object>();

        public object SelectedItem { get; set; }
    }



}
