using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordAddIn1
{
    public partial class CMERibbon
    {
        private CMEMainUserControl myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            myUserControl1 = new CMEMainUserControl();
            myUserControl1.Width = 500;
            myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(myUserControl1, "Compliance Made Easy");
            myCustomTaskPane.Width = 500;
            myCustomTaskPane.Visible = false;
        }



        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            
            myCustomTaskPane.Visible = ((RibbonToggleButton)sender).Checked;
        }

    }
}

