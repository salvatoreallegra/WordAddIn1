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
        

        public void FindAndReplaceDates()
        {
            /* var myRegex = new Regex((@"\d{2}-\d{2}-\d{4}"), RegexOptions.IgnoreCase);
             string result = myRegex.Replace("Replaced");*/

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



            object replaceAll = Word.WdReplace.wdReplaceAll;
            object ignoreCase = true;
            object wholeWord = true;
            //object forward = true;            

            if (findObject.Execute(ref missing, ref ignoreCase, ref wholeWord, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref missing, ref missing, ref missing))
            {

                object text = "Replaced " + WordToReplace + " With " + ReplacementWord;

                this.Application.ActiveDocument.Comments.Add(
                    Application.ActiveDocument.Range(), ref text);
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
