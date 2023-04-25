using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var formReviewAndAccept = new ReviewAndAccept();
           
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }



        public void DateFormatting()
        {
            //All Dates must be post Novenber 12, 2017,    *There must be a comma after the date
        }

        //Internet  Internet Internet = don't auto change    "American Internet Company"
        //, , ,, = auto change
        // comment without replace 
        public void ReplaceWithComments(string TextToFind, string ReplacementText, string CommentText)
        {        
           
            Microsoft.Office.Interop.Word.Range wordRange = null;
            Word.Document document = this.Application.ActiveDocument;

           

            //Turn off Revisions just so we won't enter the infinite loop.
            if(document.TrackRevisions == true)
            {
                document.TrackRevisions = false;
            }

            wordRange = document.Content;
            wordRange.Find.ClearFormatting();
            wordRange.Find.IgnoreSpace = true;
            wordRange.Find.Execute(FindText: TextToFind, MatchWholeWord: true,MatchWildcards: false, Forward: true);

            while (wordRange.Find.Found)
            {
                object text = CommentText;
                if (wordRange.Text == TextToFind)
                {
                    wordRange.Text = ReplacementText;
                    Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                    document.Comments.Add(
                    rng, ref text);                   
                    wordRange.Find.ClearFormatting();
                    
                }
                // Next Find
                wordRange.Find.Execute(FindText: TextToFind, MatchWholeWord:true, MatchWildcards: false, Forward: true);
            }

        }

            

            public void CommentWithoutReplace(string WordToComment, string message)
            {              
               
                Word.Document document = this.Application.ActiveDocument;
                Word.Range rng = document.Content;

                rng.Find.ClearFormatting();
                rng.Find.Forward = true;
                rng.Find.Text = WordToComment;

                rng.Find.Execute(
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);

                while (rng.Find.Found)
                {
                   
                    rng.Find.Execute(
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing);

                    object text = WordToComment + " " + message + " -CME";
                    object start = rng.Start;
                    object end = rng.End;
                                    
                    this.Application.ActiveDocument.Comments.Add(
                        Application.ActiveDocument.Range(rng.Start, rng.End), ref text);
                }

            }

            public void formatPhoneNumbers()
            {
                this.Application.ActiveDocument.Content.Select();
                Word.Find findObject = Application.Selection.Find;
                findObject.ClearFormatting();
                findObject.Text = "([(])";
                findObject.Replacement.ClearFormatting();
                findObject.Replacement.Text = "";

                object wildCard = true;
                object replaceAll = Word.WdReplace.wdReplaceAll;
                object ignoreCase = true;
                object wholeWord = true;
                //object forward = true;




                if (findObject.Execute(ref missing, ref ignoreCase, ref wholeWord, ref wildCard, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref replaceAll, ref missing, ref missing, ref missing, ref missing))
                {

                }





            }

            public void DeleteAllComments()
            {
                if (Application.ActiveDocument.Comments.Count != 0)
                {
                    this.Application.ActiveDocument.DeleteAllComments();
                    MessageBox.Show("All Comments Have Been Cleared");
                }
                else
                {
                    MessageBox.Show("There are No Comments to Delete");
                }

            }


            #region VSTO generated code

            /// <summary>
            /// Required method for Designer support - do not modify
            /// the contents of this method with the code editor.
            /// </summary>
            private void InternalStartup()
            {
                this.Startup += new System.EventHandler(ThisAddIn_Startup);
                this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            }

            #endregion
        }
    }
