using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelVerticalTab
{
    public class WindowMessageHandler : NativeWindow, IDisposable
    {
        public event EventHandler RefreshRequired;
        // 名前変更とかもろもろ飛んでくるメッセージらしい
        private static int AirSpaceNotificationMessage { get; } = (int) Helper.RegisterWindowMessage("AirSpace::Notification");

        public WindowMessageHandler(Excel.Application application)
        {
            var target = this.FindTarget(new IntPtr(application.Hwnd));
            if (target == IntPtr.Zero) return;

            AssignHandle(target);
        }

        private IntPtr FindTarget(IntPtr hwnd)
        {
            var w = Helper.FindWindowEx(hwnd, IntPtr.Zero, "XLDESK", null);
            if (w == IntPtr.Zero) return IntPtr.Zero;
            // めっちょVersion依存してそう…
            return Helper.FindWindowEx(w, IntPtr.Zero, "EXCEL7", null);
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            // Debug.WriteLine(m.Msg.ToString("X"));
            
            if (m.Msg != AirSpaceNotificationMessage) return;

            RefreshRequired?.Invoke(this, EventArgs.Empty);
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
                }

                // TODO: アンマネージ リソース (アンマネージ オブジェクト) を解放し、下のファイナライザーをオーバーライドします。
                // TODO: 大きなフィールドを null に設定します。
                ReleaseHandle();

                disposedValue = true;
            }
        }

        // TODO: 上の Dispose(bool disposing) にアンマネージ リソースを解放するコードが含まれる場合にのみ、ファイナライザーをオーバーライドします。
        ~WindowMessageHandler()
        {
            // このコードを変更しないでください。クリーンアップ コードを上の Dispose(bool disposing) に記述します。
            if (Handle != IntPtr.Zero)
                Dispose(false);
        }

        // このコードは、破棄可能なパターンを正しく実装できるように追加されました。
        public void Dispose()
        {
            // このコードを変更しないでください。クリーンアップ コードを上の Dispose(bool disposing) に記述します。
            Dispose(true);
            // TODO: 上のファイナライザーがオーバーライドされる場合は、次の行のコメントを解除してください。
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
