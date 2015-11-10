using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Configuration;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Shapes;

namespace VerticalTabControlLib
{
    public abstract class AdornerBase : Adorner, IDisposable
    {
        private AdornerLayer Layer { get; }
        protected AdornerBase(UIElement adornedElement, UIElement core) : base(adornedElement)
        {
            Core = core;

            Layer = AdornerLayer.GetAdornerLayer(adornedElement);
            Layer?.Add(this);
        }

        protected UIElement Core { get; }

        protected override Visual GetVisualChild(int index)
        {
            return Core;
        }

        protected override int VisualChildrenCount => 1;

        protected override Size MeasureOverride(Size constraint)
        {
            Core.Measure(constraint);
            return base.MeasureOverride(constraint);
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            Core.Arrange(new Rect(finalSize));
            return base.ArrangeOverride(finalSize);
        }

        #region IDisposable Support
        private bool disposedValue = false; // 重複する呼び出しを検出するには

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: マネージ状態を破棄します (マネージ オブジェクト)。
                    Layer?.Remove(this);
                }

                // TODO: アンマネージ リソース (アンマネージ オブジェクト) を解放し、下のファイナライザーをオーバーライドします。
                // TODO: 大きなフィールドを null に設定します。

                disposedValue = true;
            }
        }

        // TODO: 上の Dispose(bool disposing) にアンマネージ リソースを解放するコードが含まれる場合にのみ、ファイナライザーをオーバーライドします。
        // ~AdornerBase() {
        //   // このコードを変更しないでください。クリーンアップ コードを上の Dispose(bool disposing) に記述します。
        //   Dispose(false);
        // }

        // このコードは、破棄可能なパターンを正しく実装できるように追加されました。
        public void Dispose()
        {
            // このコードを変更しないでください。クリーンアップ コードを上の Dispose(bool disposing) に記述します。
            Dispose(true);
            // TODO: 上のファイナライザーがオーバーライドされる場合は、次の行のコメントを解除してください。
            // GC.SuppressFinalize(this);
        }
        #endregion
    }

    public class DragAdorner : AdornerBase
    {
        private const double ThisOpacity = 0.7;

        private Point StartPosition { get; set; }

        private DragAdorner(UIElement adornedElement, UIElement core) : base(adornedElement, core)
        {
        }

        public double OffsetX { get; private set; }
        public double OffsetY { get; private set; }

        public void SetOffset(double x, double y)
        {
            Debug.WriteLine($"OffsetX:{OffsetX},OffsetY:{OffsetY}");
            Debug.WriteLine($"x:{x},y:{y}");
            Debug.WriteLine($"StartX:{StartPosition.X},StartY:{StartPosition.Y}");
            OffsetX = x - StartPosition.X;
            OffsetY = y - StartPosition.Y;

            // update
            var layer = Parent as AdornerLayer;
            layer?.Update(AdornedElement);
        }

        public override GeneralTransform GetDesiredTransform(GeneralTransform transform)
        {
            var transformGroup = new GeneralTransformGroup();
            var baseObj = base.GetDesiredTransform(transform);
            if (baseObj != null)
                transformGroup.Children.Add(baseObj);

            transformGroup.Children.Add(new TranslateTransform(OffsetX, OffsetY));

            return transformGroup;
        }

        public static DragAdorner Create(UIElement adornedElement, UIElement dragTarget, Point startPosition)
        {
            // 四角形作る
            var bounds = VisualTreeHelper.GetDescendantBounds(dragTarget);
            var ghost = new Rectangle()
            {
                Width = bounds.Width,
                Height = bounds.Height,
                Fill = new VisualBrush(dragTarget) { Opacity = ThisOpacity },
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
            };

            return new DragAdorner(adornedElement, ghost)
            {
                StartPosition = startPosition,
            };
        }

    }

    public class InsertionAdorner : AdornerBase
    {
        
        private InsertionAdorner(UIElement adornedElement, UIElement core) : base(adornedElement, core)
        {
          
        }

        public InsertPositionAdorner Control { get; private set; }

        public static InsertionAdorner Create(UIElement adornedElement, bool isBottom)
        {
            var control = new InsertPositionAdorner();
            control.SetValue(HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
            control.SetValue(VerticalAlignmentProperty, isBottom ? VerticalAlignment.Bottom : VerticalAlignment.Top);

            // 何故か縮むのでサイズをセットする
            // adornedbaseのArrangeが怪しい…
            //control.Height = adornedElement.RenderSize.Height;
            //control.Width = adornedElement.RenderSize.Width;

            return new InsertionAdorner(adornedElement, control)
            {
                Control = control,
            };
        }

    }

}
