/*
- [ ] タブの同期
- [ ] リボンメニュー
- [ ] タブ移動
- [ ] コンテキストメニュー

*/

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ExcelVerticalTab.Controls;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelVerticalTab
{
    public partial class ThisAddIn
    {
        private VerticalTabHost ControlHost { get; set; }

        public ConcurrentDictionary<Excel.Workbook, WorkbookHandler> Handlers { get; } = new ConcurrentDictionary<Excel.Workbook, WorkbookHandler>(); 
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
            // ここではApplicationのイベントをハンドルしてタスクペインの生成とシート同期に紐付けるべきか

            // todo:book毎の参照を持つべし
            ControlHost = new VerticalTabHost();
            ControlHost.Initialize();
            var pane = CustomTaskPanes.Add(ControlHost, "タブ");
            pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            pane.Visible = true;
            
            this.Application.WorkbookActivate += Application_WorkbookActivate;
            
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            var handler = Handlers.GetOrAdd(Wb, x => new WorkbookHandler(x));
            // タブの同期
            handler.SyncWorksheets();
            ControlHost.AssignWorkbookHandler(handler);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
