/*
- [ ] タブの同期
- [ ] リボンメニュー
- [ ] タブ移動
- [ ] コンテキストメニュー

*/

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
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
        public ConcurrentDictionary<Excel.Workbook, PaneAndControl> Panes { get; } = new ConcurrentDictionary<Excel.Workbook, PaneAndControl>(); 

        public ConcurrentDictionary<Excel.Workbook, WorkbookHandler> Handlers { get; } = new ConcurrentDictionary<Excel.Workbook, WorkbookHandler>(); 

        private Menu RibbonMenu { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        { 
            this.Application.WorkbookActivate += Application_WorkbookActivate;
            // クローズ時の破棄をどうするか
            this.Application.WorkbookDeactivate += Application_WorkbookDeactivate;
        }

        private Tuple<Excel.Workbook, string, int> closingInfo { get; set; } = null;

        private void Application_WorkbookDeactivate(Excel.Workbook Wb)
        {
            if (this.Application.Workbooks.Count == 1)
            {
                workbookClosed(Wb);
            }
            else
            {
                closingInfo = Tuple.Create(Wb, Wb.Name, this.Application.Workbooks.Count);
            }
        }

        private PaneAndControl CreatePane(Excel.Workbook wb)
        {
            var control = new VerticalTabHost();
            control.Initialize();

            // ActiveWindowがずれてしまうことがあるっぽいような
            var w = wb.Windows.OfType<Excel.Window>().FirstOrDefault() ?? Application.ActiveWindow; // ActiveWindowは保険
            var pane = CustomTaskPanes.Add(control, "VTab", w);
            pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            pane.Visible = true;
            
            pane.VisibleChanged += Pane_VisibleChanged;

            return new PaneAndControl(pane, control);
        }

        private void Pane_VisibleChanged(object sender, EventArgs e)
        {
            RibbonMenu?.InvalidatePanesVisibility();
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            OnActivate(Wb);
        }

        public void OnActivate(Excel.Workbook wb)
        {
            if (closingInfo != null)
            {
                if (this.Application.Workbooks.Count < closingInfo.Item3)
                {
                    var res = true;
                    foreach (Excel.Workbook wb2 in this.Application.Workbooks)
                    {
                        if (wb2.Name == closingInfo.Item2)
                        {
                            res = false;
                            break;
                        }
                    }
                    if (res)
                    {
                        workbookClosed(closingInfo.Item1);
                    }
                }
                closingInfo = null;
            }

            var pane = Panes.GetOrAdd(wb, x => CreatePane(x));
            var handler = Handlers.GetOrAdd(wb, x => new WorkbookHandler(x));
            // タブの同期
            handler.SyncWorksheets();
            pane.Control.AssignWorkbookHandler(handler);

            RibbonMenu?.InvalidatePanesVisibility();
        }

        private void workbookClosed(Excel.Workbook wb)
        {
            WorkbookHandler handler = null;
            if (Handlers.TryRemove(wb, out handler))
            {
                handler.Dispose();
            }

            PaneAndControl pane = null;
            if (Panes.TryRemove(wb, out pane))
            {
                pane.Pane.VisibleChanged -= Pane_VisibleChanged;
                try
                {
                    pane.Pane.Visible = false;
                    // Windowで破棄されてるとここで死ぬ
                }
                catch (Exception)
                {
                    // Paneは死んでるし対策はいいかなって
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Application.WorkbookActivate -= Application_WorkbookActivate;

            foreach (var handler in Handlers.Values)
            {
                handler.Dispose();
            }

            Handlers.Clear();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            RibbonMenu = new Menu();
            return RibbonMenu;
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

    public class PaneAndControl
    {
        public PaneAndControl(CustomTaskPane pane, VerticalTabHost control)
        {
            Pane = pane;
            Control = control;
        }

        public CustomTaskPane Pane { get; }
        public VerticalTabHost Control { get; }
    }

}
