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
                

        private void btnClearComments_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.DeleteAllComments();
            
        }

        private void btnCorrectDocument_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.ReplaceWithComments("Internet", "internet", "Replaced Internet: Internet should not be capitalized");
            Globals.ThisAddIn.ReplaceWithComments("Intranet", "intranet", "Replaced Intranet: Intranet should not be capitalized");
            Globals.ThisAddIn.ReplaceWithComments("Web", "web", "Replaced Web: Web should not be capitalized");
            //Globals.ThisAddIn.ReplaceWithComments2("07/04/2022", "4th of July, 2022", "Changed to Proper Date Format, Nth of Month, Year");
            //Globals.ThisAddIn.ReplaceWithComments2("(702)-324-5587", "702-324-5587", "Replaced (702)-324-5587: Formated Phone Number without Parenthesis ");
            Globals.ThisAddIn.CommentWithoutReplace("cosigners", "Should Say >>> <other signatories>");
            Globals.ThisAddIn.CommentWithoutReplace("7028559999", "Phone Number should be in the format XXX-XXX-XXXX");
           

        }

       
    }
}
