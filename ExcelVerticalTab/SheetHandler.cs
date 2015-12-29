using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using ExcelVerticalTab.Annotations;
using Microsoft.Office.Interop.Excel;

namespace ExcelVerticalTab
{
    public class SheetHandler : INotifyPropertyChanged
    {
        public SheetHandler(dynamic sheet)
        {
            TargetSheet = sheet;
            Initialize();
        }

        public dynamic TargetSheet { get; }

        private string _Header;
        public string Header
        {
            get { return _Header; }
            set
            {
                if (EqualityComparer<string>.Default.Equals(_Header, value)) return;

                _Header = value;
                OnPropertyChanged();
            }
        }

        private void Initialize()
        {
            Header = TargetSheet.Name;
            
            // 名前が変更された時、っていうイベントが無いのだな…
        }

        public override string ToString()
        {
            return Header;
        }

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
