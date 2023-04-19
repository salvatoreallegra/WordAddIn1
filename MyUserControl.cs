using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Reflection;

namespace WordAddIn1
{
    public partial class MyUserControl : UserControl
    {
        public MyUserControl()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
         

            Globals.ThisAddIn.ReplaceWithComments("Internet", "internet");
            Globals.ThisAddIn.ReplaceWithComments("Intranet", "intranet");
            Globals.ThisAddIn.ReplaceWithComments("Web", "web");
            Globals.ThisAddIn.ReplaceWithComments("Website", "website");
            Globals.ThisAddIn.CommentWithoutReplace("cosigners", "Should Say >>> <other signatories>");
           // Globals.ThisAddIn.formatPhoneNumbers();
           // Globals.ThisAddIn.DateFormatting();
        }

        private void btnClearComments_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.DeleteAllComments();
        }

        private void btnCorrect_Click(object sender, EventArgs e)
        {
            //Globals.ThisAddIn.Capitalization();

            Globals.ThisAddIn.ReplaceWithComments("Internet", "internet");
            //  Globals.ThisAddIn.ReplaceWithComments("Intranet", "intranet");
            //  Globals.ThisAddIn.ReplaceWithComments("Web", "web");
            // Globals.ThisAddIn.ReplaceWithComments("Website", "website");
            //Globals.ThisAddIn.CommentWithoutReplace("cosigners");
            Globals.ThisAddIn.formatPhoneNumbers();
            Globals.ThisAddIn.DateFormatting();
        }
    }
}
