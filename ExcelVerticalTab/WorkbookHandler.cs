using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelVerticalTab.Annotations;
using Microsoft.Office.Interop.Excel;
using VerticalTabControlLib;

namespace ExcelVerticalTab
{
    public class WorkbookHandler : ITabUserControlViewModel<SheetHandler>, INotifyPropertyChanged, IDisposable
    {
        public WorkbookHandler(Workbook workbook)
        {
            TargetWorkbook = workbook;
            Initialize();
        }
        public Workbook TargetWorkbook { get; }

        public ObservableCollection<SheetHandler> Items { get; set; } = new ObservableCollection<SheetHandler>();

        #region SelectedSheet
        private SheetHandler _SelectedItem;
        public SheetHandler SelectedItem
        {
            get { return _SelectedItem; }
            set
            {
                if (EqualityComparer<SheetHandler>.Default.Equals(_SelectedItem, value)) return;

                _SelectedItem = value;
                OnPropertyChanged();
                OnSelectedSheetChanged(_SelectedItem);
            }
        }
        #endregion

        private void Initialize()
        {
            // シート追加時
            TargetWorkbook.NewSheet += TargetWorkbook_NewSheet;
            // シート削除前
            TargetWorkbook.SheetBeforeDelete += TargetWorkbook_SheetBeforeDelete;
            // シートアクティブ
            TargetWorkbook.SheetActivate += TargetWorkbook_SheetActivate;
        }

        public void SyncWorksheets()
        {
            Items.Clear();
            
            foreach (var worksheet in TargetWorkbook.Worksheets.Cast<Worksheet>())
            {
                Items.Add(new SheetHandler(worksheet));    
            }

            SelectedItem = GetSheetHandler(TargetWorkbook.ActiveSheet);
        }

        private SheetHandler GetSheetHandler(Worksheet sheet)
        {
            return Items.FirstOrDefault(x => x.TargetSheet == sheet);
        }

        private void TargetWorkbook_NewSheet(object Sh)
        {
            SyncWorksheets();
        }
        private void TargetWorkbook_SheetBeforeDelete(object Sh)
        {
            var sheet = Sh as Worksheet;
            if (sheet == null) return;

            Items.Remove(GetSheetHandler(sheet));
        }
        private void TargetWorkbook_SheetActivate(object Sh)
        {
            var sheet = Sh as Worksheet;
            if (sheet == null) return;

            SelectedItem = GetSheetHandler(sheet);
        }

        public event EventHandler<EventArgs<SheetHandler>> SelectedSheetChanged;

        protected void OnSelectedSheetChanged(SheetHandler sheetHandler)
        {
            sheetHandler?.TargetSheet.Activate();
            SelectedSheetChanged?.Invoke(this, Helper.CreateEventArgs(sheetHandler));
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

                disposedValue = true;
            }
        }

        // TODO: 上の Dispose(bool disposing) にアンマネージ リソースを解放するコードが含まれる場合にのみ、ファイナライザーをオーバーライドします。
        // ~WorkbookHandler() {
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

        #region Implements INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }
}
