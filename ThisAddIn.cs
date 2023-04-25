﻿using System;
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

        public List<Tuple<int, string, string, string, string, string>> ProcessDocument(string documentContent)
        {
            string textToSearch = documentContent;
            // Item1 = method to use, Item2 = Regex, Item3 = Find search item, Item4 = replacement, Item5 = comments, Item6 = search items            
            List<Tuple<int, string, string, string, string, string>> searchArray = BuildPatternArray();

            List<string> words = new List<string>();
            // Copy regex to an array
            foreach (var item in searchArray)
            {
                words.Add(item.Item2);

            }
            // Join all the regex in the array into a string
            string pattern = string.Join("|", words);
            // Set the regex pattern from the string
            Regex patternRegex = new Regex(pattern);

            // Get all the matches from the regex
            MatchCollection matches = patternRegex.Matches(textToSearch);
            // Create list to use to process the matches
            List<Tuple<int, string, string, string, string, string>> processArray = new List<Tuple<int, string, string, string, string, string>>();

            foreach (var match in matches)
            {
                foreach (var item in searchArray)
                {
                    Regex reg = new Regex(item.Item2);
                    Match matchedItem = reg.Match(match.ToString());
                    if (matchedItem.Success)
                    {
                        bool containsListItem = processArray.Contains(item);
                        if (!containsListItem)
                        {
                            processArray.Add(item);
                        }
                        break;
                    }

                }


            }

            return processArray;
        }

    

    private static List<Tuple<int, string, string, string, string, string>> BuildPatternArray()
    {   // Item1 = method to use, Item2 = Regex, Item3 = Find search item, Item4 = replacement, Item5 = comments, Item6 = search items
        //Method 1 - apply_changes_to_word_permutations
        //Method 2 - comment_on_change_to_word_permutations
        //Method 3 - replace_with_comments
        //Method 4 - add_comments
        //Method 5 - phone number replace
            List<Tuple<int, string, string, string, string, string>> styleArray = new List<Tuple<int, string, string, string, string, string>>();
        styleArray.Add(new Tuple<int, string, string, string, string, string>(3, "veteran", "veteran", "Veteran", "Veteran(s) should be capitalized", "True, False, False"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "Department-Wide|[dD]epartment [wW]ide|department-[wW]ide", "Department-Wide,[dD]epartment [wW]ide,department-[wW]ide", " Department-wide", "Department-wide should be capitalized and have a hyphen", "False, True, False"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(3, @"\bnation\b", "nation", "Nation", "Nation should be capitalized", "True, False, True"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(3, "congress", "congress", "Congress", "Congress / Congressional should be capitalized", "True, False, False"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(5, "[0-9]{10}", "[0-9]{10}", null, "phone number should be in the format XXX-XXX-XXXX", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "service member|[sS]ervice Member", "service member,[sS]ervice Member", " Service member", "Service member(s) should be capitalized", "False, True, False"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "members of congress|Members of congress|members of Congress", "members of congress,Members of congress,members of Congress", " Members of Congress", "Members of Congress should be capitalized", "True, False, True"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "coworkers|Coworkers|Co workers|co-workers|co workers", "coworkers,Coworkers,Co workers,co-workers,co workers", " Co-workers", " Co-workers should be capitalized", "True, False, True"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "Internet", "Internet", null, "internet should not be capitalized", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "Fiscal [yY]ear|[fF]iscal Year", "Fiscal [yY]ear", null, "fiscal year should not be capitalized.", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "Intranet", "Intranet", null, "intranet should not be capitalized", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "Web", "<Web>", null, "web should not be capitalized", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "Armed forces|armed [fF]orces", "armed [fF]orces,Armed forces", " Armed Forces", "Armed Forces should be capitalized", "False, True, False"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "[eE]-mail|Email", "[Ee]-mail,Email", null, "email should not be capitalized nor have a hyphen", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(1, @"Alzheimer's Disease|alzheimer's disease|alzheimer's Disease|Alzheimer'sdisease|alzheimer'sdisease|alzheimer'sDisease", @"Alzheimer's Disease,alzheimer's disease,alzheimer's Disease,Alzheimer'sdisease,alzheimer'sdisease,alzheimer'sDisease", @" Alzheimer's disease", "Alzheimer's should be capitalized", "True, False, True"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "governmentWide|governmentwide", "governmentWide,governmentwide", " Governmentwide", "Governmentwide should be capitalized", "False, True, False"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "Veterans integrated service network|veterans integrated service network", "Veterans integrated service network", " Veterans Integrated Service Network", " Veterans Integrated Service Network should be capitalized", "True, False, True"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "Hepatitis [cC]|hepatitis c", "Hepatitis [cC],hepatitis c", null, "hepatitis C should not be capitalized", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "Website|[wW]eb site", "Website,[wW]eb site", null, "website(s) should not be capitalized nor written as two words", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "health Care|Health [cC]are|[hH]ealthcare", "health Care,Health [cC]are,[hH]ealthcare", null, "health care should not be capitalized", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "Web[ -][bB]ased|web[ -]Based", "<Web[ -][bB]ased>,<web[ -]Based>", null, "web-based should not be capitalized", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "Posttraumatic [sS]tress [dD]isorder|posttraumatic Stress [dD]isorder|posttraumatic [sS]tress Disorder", "Posttraumatic [sS]tress [dD]isorder,posttraumatic Stress [dD]isorder,posttraumatic [sS]tress Disorder", null, "posttraumatic stress disorder should not be capitalized", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "Central office|central [oO]ffice", "Central office, central [oO]ffice", " Central Office", "Central Office should be capitalized", "False, True, False"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "federal", "federal", "Federal", "Federal should be capitalized", "True, False, True"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(1, @"\bfederal\b(?!\W+[gG]overnment\b)", "Federal government", "Federal Government", "Federal Government should be capitalized", "True, False, True"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "Veteran-Owned|[vV]eteran [oO]wned|veteran-[oO]wned", "Veteran-Owned,[vV]eteran [oO]wned,veteran-[oO]wned", " Veteran-owned", "Veteran-owned should be capitalized and have a hyphen", "True, False, True"));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(4, @"[eE]xecutive [oO]rder", "[eE]xecutive [oO]rder", null, "Executive Order should be capitalized when using EO number", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(4, @"[mM]edical [cC]enter", "[mM]edical [cC]enter", null, "Medical center should be capitalized if preceded by the formal name of the facility", null));
        styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "cosigners", "cosigners", null, "Should be other signatories", null));


        return styleArray;
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
