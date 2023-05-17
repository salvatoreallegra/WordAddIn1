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

        /*int numberOfWords = 0;
        int progressIncrement = 0;*/
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
            /*cmeTimer.Enabled = true;
            Word.Range rng = Globals.ThisAddIn.Application.ActiveDocument.Content;
            rng.Select();
            
            numberOfWords = Globals.ThisAddIn.Application.ActiveDocument.Words.Count;*/

            lblProcessingUpdates.Text = "";

            //cmeTimer.Start;
            //cmeTimer.Enabled = true;
            cmeProgress.Value = 0;
            lblProcessingUpdates.Text = "Processing, Please Be Patient and Don't Close Word";
            int progressIncrement = 0;

            var returnValue = Globals.ThisAddIn.ProcessDocument();
            foreach(var x in returnValue)
            {
                progressIncrement = 80 / returnValue.Count;
                cmeProgress.Value += progressIncrement;
                lblProcessingUpdates.Text = cmeProgress.Value + "% Complete";

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
            
            int remainingProgress = 100 - cmeProgress.Value;
            if(remainingProgress != 80)
            {
                remainingProgress = 80;
            }

           // cmeProgress.Value = remainingProgress + 10;
            lblProcessingUpdates.Text = cmeProgress.Value + "% Complete";


            Globals.ThisAddIn.FormatDate();

            cmeProgress.Value = cmeProgress.Value + 5;
            lblProcessingUpdates.Text = cmeProgress.Value + "% Complete";


            Globals.ThisAddIn.FormatNumbersUnder10();

            cmeProgress.Value = cmeProgress.Value + 5;
            lblProcessingUpdates.Text = cmeProgress.Value + "% Complete";

            Globals.ThisAddIn.DollarSymbolFollowedByDigits();
              
            cmeProgress.Value = 100;
            lblProcessingUpdates.Text = "100% complete";


            //Delete later
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

        //Probably won't end up using the timer.
        private void cmeTimer_Tick(object sender, EventArgs e)
        {
           /* if (cmeProgress.Value < 100)
            {
                cmeProgress.Value += progressIncrement + 10;
            }*/
        }
    }
}
