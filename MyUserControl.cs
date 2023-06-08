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
            Globals.ThisAddIn.DeleteAllComments(true);
            lblProcessingUpdates.Text = "";
            cmeProgress.Value = 0;
            
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

            if (Globals.ThisAddIn.Application.ActiveDocument.Revisions.Count >= 1)
            {
                Globals.ThisAddIn.Application.ActiveDocument.Revisions.AcceptAll();
            }

            Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = false;



            Globals.ThisAddIn.DeleteAllComments(false);




            lblProcessingUpdates.Text = "";

            //cmeTimer.Start;
            //cmeTimer.Enabled = true;
            cmeProgress.Value = 0;
            
            int progressIncrement = 0;

            var returnValue = Globals.ThisAddIn.ProcessDocument();
            foreach(var x in returnValue)
            {
                progressIncrement = 80 / returnValue.Count;
                cmeProgress.Value += progressIncrement;
                

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


            Globals.ThisAddIn.FormatDate();

            cmeProgress.Value = cmeProgress.Value + 5;
            Globals.ThisAddIn.FindAndReplaceSpacesAroundHyphens();
            //Globals.ThisAddIn.FindAndReplaceWildcardPlayGround("([0-9]{1,2})/([0-9]{1,2})/([0-9]{1,2})", "\\3/\\1/\\2", "Replaced with U.K. Date Format");

            //Globals.ThisAddIn.FormatNumbersUnder10();

            cmeProgress.Value = cmeProgress.Value + 5;
           

            Globals.ThisAddIn.DollarSymbolFollowedByDigits();
            
              
            cmeProgress.Value = 100;
            //lblProcessingUpdates.Text = "100% complete";
            //
            Globals.ThisAddIn.FindReplaceAndCommentWithWildCards("([0-9]{1,2})/([0-9]{1,2})/([0-9]{1,2})", "\\3/\\1/\\2", "Replaced with U.K. Date Format");
            Globals.ThisAddIn.FindAndCommentWithWildCards("Copland", "Wha a great movie");

            //Replace with Comments on Email Permutations, We can uncomment the va code just to comment
            string[] emailPermutations = new string[] { "[eE]-mail", "Email", "[Ee]-mail", "Email" };
            foreach(var email in emailPermutations)
            {
                Globals.ThisAddIn.ReplaceWithCommentsNonStyleArray(email, "email", "howdy");
            }

            //Globals.ThisAddIn.FindAndCommentWithWildCards("[1-9]?[0-9]", "Test Gimme");
            //Globals.ThisAddIn.ReplaceWithCommentsLoopThroughSentences("bill", "gem", "Howdy");
            /*  for (int i = 0; i <= 9; i++)
              {
                  Globals.ThisAddIn.IsDigitInSentence(i.ToString());
              }*/
            Globals.ThisAddIn.processSentences();
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
