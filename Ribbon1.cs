using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private MyUserControl myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            myUserControl1 = new MyUserControl();
            myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(myUserControl1, "Compliance Made Easy");
            myCustomTaskPane.Visible = false;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
           
            


            /*myUserControl1 = new MyUserControl();
               myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "Compliance Made Easy");
               myCustomTaskPane.Visible = true;*/
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            /*if (myCustomTaskPane.Visible == true)
            {
                myCustomTaskPane.Visible = false;

            }
            if (myCustomTaskPane.Visible == false)
            {
                myCustomTaskPane.Visible = true;
            }*/
            myCustomTaskPane.Visible = ((RibbonToggleButton)sender).Checked;

        }
    }
}

