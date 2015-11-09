using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

        private void CmdRefresh_OnClick(object sender, RoutedEventArgs e)
        {
            var vm = DataContext as ITabUserControlViewModel;
            vm?.Refresh_Required();
        }

        private void LstTab_OnTryMoved(object sender, TryMoveEventArgs e)
        {
            var vm = DataContext as ITabUserControlViewModel;
            vm?.TryMoved(e.Source, e.Target);
        }
    }

    public interface ITabUserControlViewModel
    {
        void Refresh_Required();

        string InputToFilter { get; set; }

        void TryMoved(object source, object target);
    }

    public interface ITabUserControlViewModel<T> : ITabUserControlViewModel
    {
        ObservableCollection<T> Items { get; set; }
        ICollectionView ItemsView { get; set; }
        T SelectedItem { get; set; }
    }

    internal class TabUserControlViewModelMock : ITabUserControlViewModel<object>
    {
        public ObservableCollection<object> Items { get; set; } = new ObservableCollection<object>();

        private ICollectionView _ItemsView;

        public ICollectionView ItemsView
        {
            get { return _ItemsView ?? (_ItemsView = CollectionViewSource.GetDefaultView(Items)); }
            set { _ItemsView = value; }
        } 

        public object SelectedItem { get; set; }

        public string InputToFilter
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public void Refresh_Required()
        {
            throw new NotImplementedException();
        }

        public void TryMoved(object source, object target)
        {
            throw new NotImplementedException();
        }
    }

}
