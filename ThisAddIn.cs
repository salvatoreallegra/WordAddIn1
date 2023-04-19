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

            

        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        
             
       
        public void DateFormatting()
        {
            //All Dates must be post Novenber 12, 2017,    *There must be a comma after the date

         
        }

        public void ReplaceWithComments(string WordToReplace, string ReplacementWord)           
        {
            

            this.Application.ActiveDocument.Content.Select();
            Word.Find findObject = Application.Selection.Find;
            
            findObject.ClearFormatting();
            findObject.Text = WordToReplace;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = ReplacementWord;
            object matchWholeWord = true;
            object matchCase = true;

            object replaceAll = Word.WdReplace.wdReplaceAll;
           
            

            findObject.Execute(ref missing, ref matchCase, ref matchWholeWord, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceAll, ref missing, ref missing, ref missing, ref missing);

            //Add comments
            Word.Document document = this.Application.ActiveDocument;
            Word.Range rng = document.Content;
            Object text = "Replaced " + WordToReplace + " With " + ReplacementWord;
            this.Application.ActiveDocument.Comments.Add(
                    Application.ActiveDocument.Range(rng.Start, rng.End), ref text);



        }



          

        public void CommentWithoutReplace(string WordToComment, string message)
        {
           
         /*   if (Application.ActiveDocument.Comments.Count != 0)
            {
                this.Application.ActiveDocument.DeleteAllComments();
                //MessageBox.Show("Re-Setting Comments prior to correction");
            } */ 
            int intFound = 0;
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
                intFound++;
                rng.Find.Execute(
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);

                object text = WordToComment + " " + message + " -CME";
                object start = rng.Start;
                object end = rng.End;

                //Word.Range commentRange = this.Range(ref start, ref end);
                

                this.Application.ActiveDocument.Comments.Add(
                    Application.ActiveDocument.Range(rng.Start,rng.End), ref text);
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

                //object text = "Replaced " + WordToReplace + " With " + ReplacementWord;

               /* this.Application.ActiveDocument.Comments.Add(
                    Application.ActiveDocument.Range(), ref text);*/
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
