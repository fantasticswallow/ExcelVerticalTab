using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using VerticalTabControlLib;
using UserControl = System.Windows.Forms.UserControl;

namespace ExcelVerticalTab.Controls
{
    public partial class VerticalTabHost : UserControl
    {
        public VerticalTabHost()
        {
            InitializeComponent();
            TabControl = new TabUserControl();
        }

        public TabUserControl TabControl { get; }

        public WorkbookHandler CurrentHandler { get; private set; }

        private void VerticalTabHost_Load(object sender, System.EventArgs e)
        {
            var host = new ElementHost()
            {
                Dock = DockStyle.Right,
                Child = TabControl,
            };
            
            // シートとの連携をどこでやるか
            
            Controls.Add(host);
        }

        public void Initialize()
        {
            
        }

        public void AssignWorkbookHandler(WorkbookHandler handler)
        {
            CurrentHandler = handler;

            TabControl.DataContext = handler;
        }

    }
}
