using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Interop;

namespace KMIQ
{
    public partial class Menu
    {
        private void Menu_Load(object sender, RibbonUIEventArgs e)
        {
            this.cb_Student.Text = Globals.ThisWorkbook.Main.PersonCombo_Selector;
        }

        private void btn_LoadData_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Main.DataLoad.GetOpenFile();
        }

        private void btn_ClearDB_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Main.DataLoad.ClearDBCheck();
        }

        private void cb_Student_ItemsLoading(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Main.MakeReport.SetMenuPersonalCombo(sender, this);
        }

        private void cb_Student_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Main.MakeReport.SelectedPersonalCombo(sender);
        }

        private void btn_PrintOut_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Main.MakeReport.ShowPrint();
            //Globals.ThisWorkbook.Main.SaveReport.SaveReport();
        }

        private void btn_About_Click(object sender, RibbonControlEventArgs e)
        {
            AboutBox1 box = new AboutBox1();
            box.ShowDialog();
        }

        private void btn_Online_Publish_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Main.MakeReport.ClearAllReport();
            Form.OnlineDataManager odm = new Form.OnlineDataManager();
            new WindowInteropHelper(odm).Owner = Globals.ThisWorkbook.Main.G_WINDOW_HANDLE;
            odm.ShowDialog();
        }
    }
}
