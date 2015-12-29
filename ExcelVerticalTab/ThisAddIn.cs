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
        public ConcurrentDictionary<string, PaneAndControl> Panes { get; } = new ConcurrentDictionary<string, PaneAndControl>(); 

        // public ConcurrentDictionary<Excel.Workbook, WorkbookHandler> Handlers { get; } = new ConcurrentDictionary<Excel.Workbook, WorkbookHandler>(); 

        private ConcurrentQueue<string> CleanQueue { get; } = new ConcurrentQueue<string>(); 

        private Menu RibbonMenu { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        { 
            this.Application.WorkbookActivate += Application_WorkbookActivate;
            // クローズ時の破棄をどうするか
            this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
            this.Application.WindowDeactivate += Application_WindowDeactivate;
        }

        private void Application_WindowDeactivate(Excel.Workbook wb, Excel.Window wn)
        {
            ProcessCleanQueue(true);
        }

        private PaneAndControl CreatePane(Excel.Workbook wb)
        {
            var control = new VerticalTabHost();
            control.Initialize();

            // ActiveWindowがずれてしまうことがあるっぽいような
            var w = wb.Windows.OfType<Excel.Window>().FirstOrDefault() ?? Application.ActiveWindow; // ActiveWindowは保険
            var pane = CustomTaskPanes.Add(control, "VTab", w);
            pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            pane.Width = 140;
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
            ProcessCleanQueue();

            var pane = Panes.GetOrAdd(wb.Name, _ => CreatePane(wb));
            var handler = pane.Control.CurrentHandler ?? new WorkbookHandler(wb);
            // var handler = Handlers.GetOrAdd(wb, x => new WorkbookHandler(x));
            // タブの同期
            handler.SyncWorksheets();
            pane.Control.AssignWorkbookHandler(handler);

            RibbonMenu?.InvalidatePanesVisibility();
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            OnBeforeClose(wb);
        }

        public void OnBeforeClose(Excel.Workbook wb)
        {
            // 掃除キューに登録
            CleanQueue.Enqueue(wb.Name);
        }

        private void ProcessCleanQueue(bool onClose = false)
        {
            var workbooks = Application.Workbooks.Cast<Excel.Workbook>().ToArray();
            while (CleanQueue.Count > 0)
            {
                var closedBookName = "";
                if (!CleanQueue.TryDequeue(out closedBookName))
                {
                    continue;
                }

                if (!onClose && workbooks.Any(x => x.Name == closedBookName))
                {
                    continue;
                }

                PaneAndControl pane_control;
                if (Panes.TryRemove(closedBookName, out pane_control))
                {
                    CleanUpPaneAndControl(pane_control);
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Application.WorkbookActivate -= Application_WorkbookActivate;
            Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            Application.WindowDeactivate -= Application_WindowDeactivate;

            foreach (var x in Panes.Values)
            {
                CleanUpPaneAndControl(x);
            }

            Panes.Clear();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Application);
        }

        private void CleanUpPaneAndControl(PaneAndControl target)
        {
            target.Control.CurrentHandler?.Dispose();
            try
            {
                target.Pane.VisibleChanged -= Pane_VisibleChanged;
                CustomTaskPanes.Remove(target.Pane);
            }
            catch (ObjectDisposedException)
            {
                // 無視
            }
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
