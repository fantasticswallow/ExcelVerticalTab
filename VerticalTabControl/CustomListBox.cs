using System;
using System.Collections.Generic;
using System.Diagnostics;
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
    /// このカスタム コントロールを XAML ファイルで使用するには、手順 1a または 1b の後、手順 2 に従います。
    ///
    /// 手順 1a) 現在のプロジェクトに存在する XAML ファイルでこのカスタム コントロールを使用する場合
    /// この XmlNamespace 属性を使用場所であるマークアップ ファイルのルート要素に
    /// 追加します:
    ///
    ///     xmlns:MyNamespace="clr-namespace:VerticalTabControl"
    ///
    ///
    /// 手順 1b) 異なるプロジェクトに存在する XAML ファイルでこのカスタム コントロールを使用する場合
    /// この XmlNamespace 属性を使用場所であるマークアップ ファイルのルート要素に
    /// 追加します:
    ///
    ///     xmlns:MyNamespace="clr-namespace:VerticalTabControl;assembly=VerticalTabControl"
    ///
    /// また、XAML ファイルのあるプロジェクトからこのプロジェクトへのプロジェクト参照を追加し、
    /// リビルドして、コンパイル エラーを防ぐ必要があります:
    ///
    ///     ソリューション エクスプローラーで対象のプロジェクトを右クリックし、
    ///     [参照の追加] の [プロジェクト] を選択してから、このプロジェクトを選択します。
    ///
    ///
    /// 手順 2)
    /// コントロールを XAML ファイルで使用します。
    ///
    ///     <MyNamespace:CustomControl1/>
    ///
    /// </summary>
    public class CustomListBox : ListBox
    {
        public static readonly DependencyProperty EnableSortByDragAndDropProperty = DependencyProperty.Register(
            "EnableSortByDragAndDrop", typeof (bool), typeof (CustomListBox), new PropertyMetadata(default(bool)));

        public bool EnableSortByDragAndDrop
        {
            get { return (bool) GetValue(EnableSortByDragAndDropProperty); }
            set { SetValue(EnableSortByDragAndDropProperty, value); }
        }

        protected override DependencyObject GetContainerForItemOverride()
        {
            return new CustomListBoxItem();
        }

        protected override bool IsItemItsOwnContainerOverride(object item)
        {
            return item is CustomListBoxItem;
        }

        private FrameworkElement _targetContainer;
        
        private Point _startPosition;

        private DragAdorner _dragAdorner;
        private InsertionAdorner _insertionAdorner;

        private FrameworkElement GetContainer(FrameworkElement originalSource)
        {
            if (originalSource == null) return null;
            return ContainerFromElement(originalSource) as FrameworkElement;
        }

        // Arrow keys don't work after programmatically setting ListView.SelectedItem
        // https://stackoverflow.com/questions/7363777/arrow-keys-dont-work-after-programmatically-setting-listview-selecteditem
        protected override void OnSelectionChanged(SelectionChangedEventArgs e)
        {
            base.OnSelectionChanged(e);

            var container = (UIElement)ItemContainerGenerator.ContainerFromItem(SelectedItem);

            if (container != null)
            {
                container.Focus();
            }
        }

        // 左クリック前
        protected override void OnPreviewMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            base.OnPreviewMouseLeftButtonDown(e);
            if (!EnableSortByDragAndDrop) return;

            _targetContainer = GetContainer(e.OriginalSource as FrameworkElement);
            if (_targetContainer == null) return;
            
            _startPosition = PointToScreen(e.GetPosition(_targetContainer));
            //Debug.WriteLine($"StartPosition:x={_startPosition.X},y={_startPosition.Y}");
        }

        protected override void OnPreviewMouseMove(MouseEventArgs e)
        {
            base.OnPreviewMouseMove(e);
            if (!EnableSortByDragAndDrop) return;
            if (_targetContainer?.DataContext == null) return;

            var listBoxItem = _targetContainer as ListBoxItem;
            if (listBoxItem != null && !listBoxItem.IsSelected)
                return;

            // 移動量が十分か検証する
            var currentPosition = PointToScreen(e.GetPosition(_targetContainer));
            var delta = (_startPosition - currentPosition);
            if (!delta.IsEnoughMoveForDrug()) return;

            //Debug.WriteLine("DragDropStart");

            _dragAdorner = _dragAdorner ?? (_dragAdorner = DragAdorner.Create(this, _targetContainer, _startPosition));

            _dragAdorner.SetOffset(currentPosition.X, currentPosition.Y);

            DragDrop.DoDragDrop(this, _targetContainer.DataContext, DragDropEffects.Move);

            // 終わったら後始末
            ResetDragAndDropParameter();

        }

        protected override void OnPreviewMouseUp(MouseButtonEventArgs e)
        {
            base.OnPreviewMouseUp(e);

            //Debug.WriteLine("MouseUp");

            ResetDragAndDropParameter();
        }

        private void ResetDragAndDropParameter()
        {
            //Debug.WriteLine("ResetParameter");

            _targetContainer = null;
            _startPosition = new Point();

            _dragAdorner?.Dispose();
            _dragAdorner = null;

            _insertionAdorner?.Dispose();
            _insertionAdorner = null;

        }

        protected override void OnPreviewDragEnter(DragEventArgs e)
        {
            base.OnPreviewDragEnter(e);

            //Debug.WriteLine("PreviewDragEnter");

            // のっかったコンテナを取得
            var isBottom = false;
            var entered = GetContainer(e.OriginalSource as FrameworkElement);
            if (entered == null)
            {
                // 最後にしてみる
                entered = ItemContainerGenerator.ContainerFromIndex(Items.Count - 1) as FrameworkElement;
                isBottom = true;
            }

            _insertionAdorner = InsertionAdorner.Create(entered, isBottom);

        }

        protected override void OnPreviewDragOver(DragEventArgs e)
        {
            base.OnPreviewDragOver(e);

            //Debug.WriteLine("PreviewDragOver");

            var currentPosition = PointToScreen(e.GetPosition(this));
            _dragAdorner.SetOffset(currentPosition.X, currentPosition.Y);
        }

        protected override void OnPreviewDragLeave(DragEventArgs e)
        {
            base.OnPreviewDragLeave(e);

            //Debug.WriteLine("PreviewDragLeave");

            _insertionAdorner?.Dispose();
            _insertionAdorner = null;
        }

        protected override void OnPreviewDrop(DragEventArgs e)
        {
            base.OnPreviewDrop(e);

            // ドロップされた位置のアイテム
            var dropped = GetContainer(e.OriginalSource as FrameworkElement)?.DataContext;
            OnTryMoved(_targetContainer?.DataContext, dropped);
        }

        protected void OnTryMoved(object source, object target)
        {
            if (source == target) return;
            TryMoved?.Invoke(this, new TryMoveEventArgs(source, target));
        }

        public event EventHandler<TryMoveEventArgs> TryMoved;
    }

    public class CustomListBoxItem : ListBoxItem
    {
        protected override void OnMouseEnter(MouseEventArgs e)
        {
            var parant = ItemsControl.ItemsControlFromItemContainer(this);
            if (parant.IsMouseCaptured)
                parant.ReleaseMouseCapture();

            base.OnMouseEnter(e);
        }
    }

    public static class Helper
    {
        public static bool IsEnoughMoveForDrug(this Vector delta)
        {
            return Math.Abs(delta.X) > SystemParameters.MinimumHorizontalDragDistance ||
                   Math.Abs(delta.Y) > SystemParameters.MinimumVerticalDragDistance;
        }
    }

    public class TryMoveEventArgs : EventArgs
    {
        public TryMoveEventArgs(object source, object target)
        {
            Source = source;
            Target = target;
        }

        public object Source { get; }
        public object Target { get; }

    }
}
