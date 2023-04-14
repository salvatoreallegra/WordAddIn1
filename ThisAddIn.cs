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
        public void findAndReplace()
        {
            /* int intFound = 0;
             Word.Document document = this.Application.ActiveDocument;
             Word.Range rng = document.Content;

             rng.Find.ClearFormatting();
             rng.Find.Forward = true;
             //object dateText = (@"\d{2}-\d{2}-\d{4}");
             rng.Find.Text = "Find Me";

             rng.Find.Execute(
                 ref missing, ref missing, ref missing, ref missing, ref missing,
                 ref missing, ref missing, ref missing, ref missing, ref missing,
                 ref missing, ref missing, ref missing, ref missing, ref missing);*/


            //Text can be a regular expression here 
            // object text = "Comment here";

            //Add Comments to each Text

            /*while (rng.Find.Found)
            {
                intFound++;
                rng.Find.Execute(
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);

                this.Application.ActiveDocument.Comments.Add(
                            Application.ActiveDocument.Range(), ref text);
            }

            MessageBox.Show("Strings found: " + intFound.ToString());
*/



            /*     this.Application.ActiveDocument.Content.Select();
                 Word.Find findObject = Application.Selection.Find;
                 findObject.ClearFormatting();
                 findObject.Text = "find me";
                 findObject.Replacement.ClearFormatting();
                 findObject.Replacement.Text = "Found";
                 string x = Application.ActiveDocument.Range().Text;

                 object text = "Add a comment to the first paragraph.";
                 this.Application.ActiveDocument.Comments.Add(
                     Application.ActiveDocument.Range(), ref text);

                 object replaceAll = Word.WdReplace.wdReplaceAll;
                 findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                     ref missing, ref missing, ref missing, ref missing, ref missing,
                     ref replaceAll, ref missing, ref missing, ref missing, ref missing);*/

        }

        public void FindAndReplaceDates()
        {
            /* var myRegex = new Regex((@"\d{2}-\d{2}-\d{4}"), RegexOptions.IgnoreCase);
             string result = myRegex.Replace("Replaced");*/

        }
        public void Capitalization()
        {
            var capitalizationWords = new Dictionary<string, string>
            {
                {"Internet", "internet" },
                {"Intranet" ,"intranet" },
                {"Web", "web" },
                {"Website", "website" },
                {"BlackList","henderson" }
            };

            /*foreach (var capitliaztionWord in capitalizationWords)
                {
                    Word.Document document = this.Application.ActiveDocument;
                    Word.Range rng = document.Content;
                    rng.Find.ClearFormatting();
                    rng.Find.Forward = true;
                    this.Application.ActiveDocument.Content.Select();
                    Word.Find findObject = Application.Selection.Find;
                    findObject.ClearFormatting();

                    var textToChange = findObject.Text = capitliaztionWord.Key;
                    findObject.Replacement.ClearFormatting();
                    var textChangedTo = findObject.Replacement.Text = capitliaztionWord.Value;
                    string x = Application.ActiveDocument.Range().Text;

                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref missing, ref missing, ref missing);

                    object text = "Replaced " + textToChange + " with " + textChangedTo;

                    this.Application.ActiveDocument.Comments.Add(
                        Application.ActiveDocument.Range(), ref text);

                }*/


            foreach (var capitliaztionWord in capitalizationWords)
            {



                this.Application.ActiveDocument.Content.Select();
                Word.Find findObject = Application.Selection.Find;
                findObject.ClearFormatting();
                findObject.Text = capitliaztionWord.Key;     //"Internet";
                findObject.Replacement.ClearFormatting();
                findObject.Replacement.Text = capitliaztionWord.Value;       //"giblet";

                object replaceAll = Word.WdReplace.wdReplaceAll;
                object ignoreCase = true;
                object wholeWord = true;
                //object forward = true;

                //object text = "Corrected Text";
               /* this.Application.ActiveDocument.Comments.Add(
                    Application.ActiveDocument.Range(), ref text);*/


                findObject.Execute(ref missing, ref ignoreCase, ref wholeWord, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref missing, ref missing, ref missing);



            }
            

        }

        public void DateFormatting()
        {
            //All Dates must be post Novenber 12, 2017,    *There must be a comma after the date

            //[A-Z,a - z][th, st, nd, rd] of " + monthsArray(i), "the date should be written as DD(st / nd / rd / th) of " + monthsArray(i) + "(e.g. 11th of// November)"

           string dateToLookFor = "Fourth of July";
            var myRegex = new Regex((@"\w of July"), RegexOptions.IgnoreCase);

            string result = myRegex.Match(dateToLookFor).ToString();

            MessageBox.Show(result);
        }

        public void DeleteAllComments()
        {

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
