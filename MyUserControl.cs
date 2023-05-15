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
            /*Basic Process of Find and Replace/Comment from the Hacker Competition/VBA/Marjorie
             * First step is process document function, this will return a list of tuples,
             * the params of process document are....It searches the whole document and returns
             * ONLY what needs to be processed.   ProcessArray calls build search array, it builds
             * the tuple of the things we are looking for, this is where you add more search criteria
             * .  Process array takes 
             * 
             *  P
             */

            cmeProgress.Value = 0;

            var returnValue = Globals.ThisAddIn.ProcessDocument();
            foreach(var x in returnValue)
            {
                switch (x.Item1)
                {
                    case 1:
                        Globals.ThisAddIn.apply_changes_to_word_permutations(x.Item3, x.Item4, x.Item5, x.Item6);
                        break;
                    case 2:
                        Globals.ThisAddIn.comment_changes_to_word_permutations(x.Item3, x.Item4, x.Item5, x.Item6);
                        break;
                    case 3:
                        Globals.ThisAddIn.ReplaceWithComments(x.Item3,x.Item4,x.Item5, x.Item6);
                        break;
                    case 4:
                        Globals.ThisAddIn.AddComments(x.Item3, x.Item4, x.Item5, x.Item6);
                        break;
                    default:                        
                        break;
                       
                }

            }
            cmeProgress.Value = 10;

            Globals.ThisAddIn.FormatDate();

            cmeProgress.Value = 50;
            Globals.ThisAddIn.FormatNumbersUnder10();

            Globals.ThisAddIn.DollarSymbolFollowedByDigits();

            cmeProgress.Value = 100;
            //Globals.ThisAddIn.ReplaceWithComments("Internet", "internet", "Replaced Internet: Internet should not be capitalized");
            //Globals.ThisAddIn.ReplaceWithComments("Intranet", "intranet", "Replaced Intranet: Intranet should not be capitalized");
            //Globals.ThisAddIn.ReplaceWithComments("Web", "web", "Replaced Web: Web should not be capitalized");
            //Globals.ThisAddIn.ReplaceWithComments2("07/04/2022", "4th of July, 2022", "Changed to Proper Date Format, Nth of Month, Year");
            //Globals.ThisAddIn.ReplaceWithComments2("(702)-324-5587", "702-324-5587", "Replaced (702)-324-5587: Formated Phone Number without Parenthesis ");
            //Globals.ThisAddIn.CommentWithoutReplace("cosigners", "Should Say >>> <other signatories>");
            //Globals.ThisAddIn.CommentWithoutReplace("7028559999", "Phone Number should be in the format XXX-XXX-XXXX");
           

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {
            
        }
    }
}
