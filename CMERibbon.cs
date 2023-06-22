using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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

        private void btnLoadComments_Click(object sender, RibbonControlEventArgs e)
        {
            var formComments = new Form();
            formComments.Show();
        }

        /*  private void btnProcessComments_Click(object sender, RibbonControlEventArgs e)
          {

              var comments = Globals.ThisAddIn.Application.ActiveDocument.Comments;
              for(int i = 1; i <= comments.Count; i++)
              {*/

        /* RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();

         item.Label = comments[i].Range.Text;
         commentsDropDown.Items.Add(item);*/
        //MessageBox.Show(comments[i].Range.Text);
        // }

        /*foreach ( var comment in comments)
        {

            RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();

            item.Label = comment.ToString();
            dropDown1.Items.Add(item);
        }*/
        //  }


    }
}

