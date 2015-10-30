using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Data;
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
            MsgHandler = new WindowMessageHandler(workbook.Application);
            Initialize();
        }
        public Workbook TargetWorkbook { get; }

        public WindowMessageHandler MsgHandler { get; }

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

        #region ItemsView
        private ICollectionView _itemsView;
        public ICollectionView ItemsView
        {
            get { return _itemsView ?? (_itemsView = CollectionViewSource.GetDefaultView(Items)); }
            set { _itemsView = value; }
        }
        #endregion

        #region InputToFilter
        private string _inputToFilter;
        public string InputToFilter
        {
            get { return _inputToFilter; }

            set
            {
                if (EqualityComparer<string>.Default.Equals(_inputToFilter, value)) return;

                _inputToFilter = value;
                OnPropertyChanged();
                ExecuteFilter();
            }
        }
        #endregion

        private void ExecuteFilter()
        {
            Regex r = null;
            if (string.IsNullOrWhiteSpace(InputToFilter))
            {
                ItemsView.Filter = null;
                ItemsView.Refresh();
                SelectFirst();
                return;
            }

            // 一旦保留
            //using (var m = Migemo.GetDefault())
            //{
            //    Debug.WriteLine(m.Query(InputToFilter));
            //    try
            //    {
            //        r = m.GetRegex(InputToFilter);
            //    }
            //    catch (Exception e)
            //    {
            //        Debug.WriteLine(e.Message);
            //        r = null;
            //    }
                
            //}
            var comparer = StringComparer.Create(CultureInfo.CurrentCulture, true);
            
            Predicate<object> filter;
            if (r == null)
            {
                filter = x =>
                {
                    var s = x as SheetHandler;
                    var source = s?.Header ?? "";
                    const CompareOptions options = CompareOptions.IgnoreCase | CompareOptions.IgnoreKanaType | CompareOptions.IgnoreNonSpace | CompareOptions.IgnoreSymbols | CompareOptions.IgnoreWidth;
                    return CultureInfo.CurrentCulture.CompareInfo.IndexOf(source, InputToFilter, options) >= 0;
                };
            }
            else
            {
                filter = x =>
                {
                    var s = x as SheetHandler;
                    return r.IsMatch(s?.Header ?? "");
                };
            }
            
            ItemsView.Filter = filter;
            
            SelectFirst();
            
        }

        private void SelectFirst()
        {
            if (SelectedItem != null) return;
            
            var first = ItemsView.OfType<SheetHandler>().FirstOrDefault();
            if (first != null)
                SelectedItem = first;
            
        }

        private void Initialize()
        {
            // シート追加時
            TargetWorkbook.NewSheet += TargetWorkbook_NewSheet;
            // シート削除前
            TargetWorkbook.SheetBeforeDelete += TargetWorkbook_SheetBeforeDelete;
            // シートアクティブ
            TargetWorkbook.SheetActivate += TargetWorkbook_SheetActivate;

            MsgHandler.RefreshRequired += MsgHandlerOnRefreshRequired;
        }

        private void RemoveHandler()
        {
            // シート追加時
            TargetWorkbook.NewSheet -= TargetWorkbook_NewSheet;
            // シート削除前
            TargetWorkbook.SheetBeforeDelete -= TargetWorkbook_SheetBeforeDelete;
            // シートアクティブ
            TargetWorkbook.SheetActivate -= TargetWorkbook_SheetActivate;

            MsgHandler.RefreshRequired -= MsgHandlerOnRefreshRequired;

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

            Debug.WriteLine($"{sheet.Name} is Activate");

            SelectedItem = GetSheetHandler(sheet);
        }

        private void MsgHandlerOnRefreshRequired(object sender, EventArgs eventArgs)
        {
            _suppressChanged = true;
            SyncWorksheets();
            _suppressChanged = false;

        }

        public void Refresh_Required()
        {
            SyncWorksheets();
        }

        public event EventHandler<EventArgs<SheetHandler>> SelectedSheetChanged;

        private bool _suppressChanged;

        protected void OnSelectedSheetChanged(SheetHandler sheetHandler)
        {
            if (_suppressChanged) return;
            Debug.WriteLineIf(sheetHandler != null, $"{sheetHandler?.TargetSheet.Name} is Selected");
            //var location = Assembly.GetExecutingAssembly().Location;
            //Debug.WriteLine(location);
            sheetHandler?.TargetSheet.Activate();
            SelectedSheetChanged?.Invoke(this, Helper.CreateEventArgs(sheetHandler));
            //var sheet = sheetHandler?.TargetSheet;
            //if (sheet == null) return;
            //Range r = sheet.Cells[1, 1];
            //r.Value = location;
        }

        #region IDisposable Support
        private bool disposedValue = false; // 重複する呼び出しを検出するには

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    RemoveHandler();
                }

                // TODO: アンマネージ リソース (アンマネージ オブジェクト) を解放し、下のファイナライザーをオーバーライドします。
                // TODO: 大きなフィールドを null に設定します。
                MsgHandler.Dispose();

                disposedValue = true;
            }
        }

        // TODO: 上の Dispose(bool disposing) にアンマネージ リソースを解放するコードが含まれる場合にのみ、ファイナライザーをオーバーライドします。
        ~WorkbookHandler()
        {
            // このコードを変更しないでください。クリーンアップ コードを上の Dispose(bool disposing) に記述します。
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
