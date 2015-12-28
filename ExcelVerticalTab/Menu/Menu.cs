using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;

// TODO:  リボン (XML) アイテムを有効にするには、次の手順に従います。

// 1: 次のコード ブロックを ThisAddin、ThisWorkbook、ThisDocument のいずれかのクラスにコピーします。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Menu();
//  }

// 2. ボタンのクリックなど、ユーザーの操作を処理するためのコールバック メソッドを、このクラスの
//    "リボンのコールバック" 領域に作成します。メモ: このリボンがリボン デザイナーからエクスポートされたものである場合は、
//    イベント ハンドラー内のコードをコールバック メソッドに移動し、リボン拡張機能 (RibbonX) のプログラミング モデルで
//    動作するように、コードを変更します。

// 3. リボン XML ファイルのコントロール タグに、コードで適切なコールバック メソッドを識別するための属性を割り当てます。  

// 詳細については、Visual Studio Tools for Office ヘルプにあるリボン XML のドキュメントを参照してください。


namespace ExcelVerticalTab
{
    [ComVisible(true)]
    public class Menu : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Menu()
        {
        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelVerticalTab.Menu.Menu.xml");
        }

        #endregion

        #region リボンのコールバック
        //ここにコールバック メソッドを作成します。コールバック メソッドの追加方法の詳細については、http://go.microsoft.com/fwlink/?LinkID=271226 にアクセスしてください。

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void chkVisibility_Changed(Office.IRibbonControl control, bool isHide)
        {
            var book = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (book == null) return;

            var pane = Globals.ThisAddIn.Panes.GetValueOrDefault(book.Name);
            if (pane == null) return;
            
            pane.Pane.Visible = !isHide;
        }

        public void cmdRefresh_Click(Office.IRibbonControl control)
        {
            var book = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (book == null) return;

            Globals.ThisAddIn.OnActivate(book);            
        }

        public bool GetPanesVisibility(Office.IRibbonControl control)
        {
            var book = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (book == null) return false;

            var pane = Globals.ThisAddIn.Panes.GetValueOrDefault(book.Name);
            if (pane == null) return false;

            return !pane.Pane.Visible;
        }

        public void InvalidatePanesVisibility()
        {
            ribbon.InvalidateControl("chkVisibility");
        }

        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
