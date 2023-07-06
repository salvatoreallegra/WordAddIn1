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

            formProgress form = new formProgress();
            form.Show();
            ProgressBar cmeProgress = (ProgressBar)Application.OpenForms["formProgress"].Controls["cmeProgress"];

            
            
                
            
            
                //TextBox t = Application.OpenForms["Form1"].Controls["textBox1"] as TextBox;

            if (Globals.ThisAddIn.Application.ActiveDocument.Revisions.Count >= 1)
            {
                Globals.ThisAddIn.Application.ActiveDocument.Revisions.AcceptAll();
            }

            Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = false;


            Globals.ThisAddIn.DeleteAllComments(false);


            //lblProcessingUpdates.Text = "";

            //cmeTimer.Start;
            //cmeTimer.Enabled = true;
            //cmeProgress.Value = 0;

            //int progressIncrement = 0;

            var returnValue = Globals.ThisAddIn.ProcessDocument();
            foreach (var x in returnValue)
            {
               // progressIncrement = 80 / returnValue.Count;
                //cmeProgress.Value += progressIncrement;


                switch (x.Item1)
                {
                    case 1:
                        Globals.ThisAddIn.apply_changes_to_word_permutations(x.Item3, x.Item4, x.Item5, x.Item6);
                        break;
                    case 2:
                        Globals.ThisAddIn.comment_changes_to_word_permutations(x.Item2, x.Item3, x.Item4, x.Item5, x.Item6);
                        break;
                    case 3:
                        Globals.ThisAddIn.ReplaceWithComments(x.Item3, x.Item4, x.Item5, x.Item6);
                        break;
                    case 4:
                        Globals.ThisAddIn.AddComments(x.Item2, x.Item3, x.Item4, x.Item5, x.Item6);
                        break;

                    case 5:
                        Globals.ThisAddIn.ReplaceWithCommentsWholeWord(x.Item3, x.Item4, x.Item5, x.Item6);
                        break;
                    default:
                        break;

                }

            }

           /* int remainingProgress = 100 - cmeProgress.Value;
            if (remainingProgress != 80)
            {
                remainingProgress = 80;
            }
*/

            Globals.ThisAddIn.FormatDate();

            cmeProgress.Value = cmeProgress.Value += 40;
            Globals.ThisAddIn.FindAndReplaceSpacesAroundHyphens();
            //Globals.ThisAddIn.FindAndReplaceWildcardPlayGround("([0-9]{1,2})/([0-9]{1,2})/([0-9]{1,2})", "\\3/\\1/\\2", "Replaced with U.K. Date Format");

            //Globals.ThisAddIn.FormatNumbersUnder10();

            cmeProgress.Value = cmeProgress.Value + 5;


            Globals.ThisAddIn.DollarSymbolFollowedByDigits();


            //cmeProgress.Value = 100;
            //lblProcessingUpdates.Text = "100% complete";
            //
            //Globals.ThisAddIn.FindReplaceAndCommentWithWildCards("([0-9]{1,2})/([0-9]{1,2})/([0-9]{1,2})", "\\3/\\1/\\2", "Replaced with U.K. Date Format");

            /*****************************************************************************
             * Replace with Comments on Email Permutations, We can uncomment the va code just 
             * to comment without replace
             * 
             * 
             * *****************************/
            string[] emailPermutations = new string[] { "Email", "[Ee]-mail", "[eE]-Mail" };
            foreach (var email in emailPermutations)
            {
                Globals.ThisAddIn.ReplaceWithCommentsNonStyleArray(email, "email", "email should not be capitalized nor have a hyphen");
            }

            cmeProgress.Value += 10;

            /********************************
             * This function call does the
             * numbers under 10 style rule
             * *****************************/
            Globals.ThisAddIn.processSentences();
            cmeProgress.Value += 10;


            Globals.ThisAddIn.replaceVeteranInstances();
            cmeProgress.Value += 10;

            Globals.ThisAddIn.replaceFederalInstances();
            cmeProgress.Value += 10;

            Globals.ThisAddIn.replaceCongressInstances();
            cmeProgress.Value += 10;

            Globals.ThisAddIn.commentWebInstances();
            Globals.ThisAddIn.commentCentralOffice();
            Globals.ThisAddIn.commentCoWorkers();
            cmeProgress.Value += 5;


            form.Close();

            MessageBox.Show("Document Processing Complete");

            //myCustomTaskPane.Visible = ((RibbonToggleButton)sender).Checked;
        }

        private void btnClearComments_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveDocument.Comments.Count != 0)
            {
                Globals.ThisAddIn.Application.ActiveDocument.DeleteAllComments();
              //  if (showMessageBox)
                //{
                    // ShowAllComments();
                    MessageBox.Show("Clearing All Comments ");
              //  }
            }
            else
            {
               // if (showMessageBox)
              //  {
                    MessageBox.Show("There are No Comments to Delete");
              //  }

            }

            //Globals.ThisAddIn.DeleteAllComments(true);

        }
    }
}

