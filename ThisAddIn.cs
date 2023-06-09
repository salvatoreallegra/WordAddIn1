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
using Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private int numberCounter = 0;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        /******************************************
         * This is going to be the main Find and Replace
         * with Wildcards function for the entire app.        
         *
         * ****************************************/

        public void FindReplaceAndCommentWithWildCards(string wildCardText, string replacementText, string commentMessage)
        {
            //do this, it only takes one step when undo
            Microsoft.Office.Interop.Word.UndoRecord ur = this.Application.UndoRecord;
            ur.StartCustomRecord("FindReplaceAndCommentWithWildCards");

            Microsoft.Office.Interop.Word.Range wordRange = null;
            Word.Document document = this.Application.ActiveDocument;
            wordRange = document.Content;
            wordRange.Find.ClearFormatting();
            wordRange.Find.ClearAllFuzzyOptions();
            wordRange.Find.Replacement.ClearFormatting();
            wordRange.Find.IgnoreSpace = true;
            wordRange.Find.MatchCase = false;
            //wordRange.Find.MatchWildcards = true;
            wordRange.Find.Text = wildCardText;

            wordRange.Find.Replacement.Text = replacementText;//that's the right way to write
            wordRange.Find.Forward = true;
            wordRange.Find.Wrap = WdFindWrap.wdFindStop;

            //don't forget the Replace argument
            wordRange.Find.Execute(MatchWildcards: true, Replace: WdReplace.wdReplaceOne);//Just set the argument MatchWildcards here!! like you wrote in this line : wordRange.Find.Execute(FindText: wildCardText, MatchCase: false, MatchWildcards: true);
            while (wordRange.Find.Found)
            {
                object commentText = "[REPLACED] " + commentMessage + " -CME";
                Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                //rng.Text = replacementText;//This is wrong!! refer to above
                document.Comments.Add(rng, ref commentText);
                wordRange.Find.ClearFormatting();

                // Next Find
                //don't forget to reset the range wordRange
                wordRange.SetRange(wordRange.End, wordRange.Document.Content.End);

                //wordRange.Find.Execute(FindText: wildCardText, MatchCase: false, MatchWildcards: true);
                wordRange.Find.Execute(MatchWildcards: true, Replace: WdReplace.wdReplaceOne);
            }

            ur.EndCustomRecord();
        }

        /******************************************
         * This is going to be the main Find and Comment
         * with Wildcards function for the entire app.
         * There is no replace here, just a wildcard find
         * and adding a comment to the found range
         * ****************************************/

        public void FindAndCommentWithWildCards(string wildCardText, string commentMessage)
        {
            //do this, it only takes one step when undo
            Microsoft.Office.Interop.Word.UndoRecord ur = this.Application.UndoRecord;
            ur.StartCustomRecord("FindReplaceAndCommentWithWildCards");

            Microsoft.Office.Interop.Word.Range wordRange = null;
            Word.Document document = this.Application.ActiveDocument;
            wordRange = document.Content;
            wordRange.Find.ClearFormatting();
            wordRange.Find.ClearAllFuzzyOptions();
            wordRange.Find.Replacement.ClearFormatting();
            wordRange.Find.IgnoreSpace = true;
            wordRange.Find.MatchCase = false;
            //wordRange.Find.MatchWholeWord = true;
            //wordRange.Find.MatchWildcards = true;
            wordRange.Find.Text = wildCardText;

            //wordRange.Find.Replacement.Text = replacementText;//that's the right way to write
            wordRange.Find.Forward = true;
            wordRange.Find.Wrap = WdFindWrap.wdFindStop;

            //don't forget the Replace argument
            wordRange.Find.Execute(MatchWildcards: true );//Just set the argument MatchWildcards here!! like you wrote in this line : wordRange.Find.Execute(FindText: wildCardText, MatchCase: false, MatchWildcards: true);
            while (wordRange.Find.Found)
            {
                object commentText = commentMessage;
                Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                //rng.Text = replacementText;//This is wrong!! refer to above
                document.Comments.Add(rng, ref commentText);
                wordRange.Find.ClearFormatting();

                // Next Find
                //don't forget to reset the range wordRange
                wordRange.SetRange(wordRange.End, wordRange.Document.Content.End);

                //wordRange.Find.Execute(FindText: wildCardText, MatchCase: false, MatchWildcards: true);
                wordRange.Find.Execute(MatchWildcards: true);
            }

            ur.EndCustomRecord();
        }

        /******************************************
         * We are using this for the web/website etc...rule
         * with Wildcards function for the entire app.
         * There is no replace here, just a wildcard find
         * and adding a comment to the found range
         * ****************************************/

        public void FindAndCommentWithWholeWord(string wildCardText, string commentMessage)
        {
            //do this, it only takes one step when undo
            Microsoft.Office.Interop.Word.UndoRecord ur = this.Application.UndoRecord;
            ur.StartCustomRecord("FindReplaceAndCommentWithWildCards");

            Microsoft.Office.Interop.Word.Range wordRange = null;
            Word.Document document = this.Application.ActiveDocument;
            wordRange = document.Content;
            wordRange.Find.ClearFormatting();
            wordRange.Find.ClearAllFuzzyOptions();
            wordRange.Find.Replacement.ClearFormatting();
            wordRange.Find.IgnoreSpace = true;
            wordRange.Find.MatchCase = false;
            wordRange.Find.MatchWholeWord = true;
            //wordRange.Find.MatchWildcards = true;
            wordRange.Find.Text = wildCardText;

            //wordRange.Find.Replacement.Text = replacementText;//that's the right way to write
            wordRange.Find.Forward = true;
            wordRange.Find.Wrap = WdFindWrap.wdFindStop;

            //don't forget the Replace argument
            wordRange.Find.Execute(MatchWildcards: false,MatchWholeWord:true);//Just set the argument MatchWildcards here!! like you wrote in this line : wordRange.Find.Execute(FindText: wildCardText, MatchCase: false, MatchWildcards: true);
            while (wordRange.Find.Found)
            {
                object commentText = commentMessage;
                Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                //rng.Text = replacementText;//This is wrong!! refer to above
                document.Comments.Add(rng, ref commentText);
                wordRange.Find.ClearFormatting();

                // Next Find
                //don't forget to reset the range wordRange
                wordRange.SetRange(wordRange.End, wordRange.Document.Content.End);

                //wordRange.Find.Execute(FindText: wildCardText, MatchCase: false, MatchWildcards: true);
                wordRange.Find.Execute(MatchWildcards: true);
            }

            ur.EndCustomRecord();
        }

        /******************************************
        * Main Find and Comment
        * with Wildcards function for the entire app
        * that searches inside a sentence only, instead of the entire
        * document.
        * There is no replace here, just a wildcard find
        * and adding a comment to the found range
        * ****************************************/

        public void FindAndCommentInSentence(string wildCardText, string comment)
        {
            //do this, it only takes one step when undo
            Microsoft.Office.Interop.Word.UndoRecord ur = this.Application.UndoRecord;
            ur.StartCustomRecord("FindAndCommentInSentence");

            Microsoft.Office.Interop.Word.Range wordRange = null;
            Word.Document document = this.Application.ActiveDocument;

            var sentenceCount = document.Sentences.Count;
            for (int i = 1; i <= sentenceCount; i++)
            {

                wordRange = this.Application.ActiveDocument.Sentences[i];
                //wordRange.Bold = 1;
                wordRange.Find.ClearFormatting();
                wordRange.Find.ClearAllFuzzyOptions();
                wordRange.Find.Replacement.ClearFormatting();
                wordRange.Find.IgnoreSpace = true;
                wordRange.Find.MatchCase = false;
                wordRange.Find.MatchWildcards = true;
                wordRange.Find.Text = wildCardText;
               // wordRange.Find.Replacement.Text = replacementText;//that's the right way to write
                wordRange.Find.Forward = true;
                wordRange.Find.Wrap = WdFindWrap.wdFindStop;

                //don't forget the Replace argument
                wordRange.Find.Execute(MatchWildcards: true);//Just set the argument MatchWildcards here!! like you wrote in this line : wordRange.Find.Execute(FindText: wildCardText, MatchCase: false, MatchWildcards: true);
                while (wordRange.Find.Found)
                {
                   // MessageBox.Show(wordRange.Text);

                  //   bool foundZeroThroughNime = true;
                   

                    object commentText = comment;
                    Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                    //rng.Text = replacementText;//This is wrong!! refer to above
                    document.Comments.Add(rng, ref commentText);
                    wordRange.Find.ClearFormatting();


                    // Next Find
                    //don't forget to reset the range wordRange
                    wordRange.SetRange(wordRange.End, wordRange.Document.Content.End);
                    //wordRange.SetRange(wordRange.Sentences[i]);

                    //wordRange.Find.Execute(FindText: wildCardText, MatchCase: false, MatchWildcards: true);
                       wordRange.Find.Execute(MatchWildcards: true);
                }
            }

            ur.EndCustomRecord();
        }

        /**********************************
         * 
         * 
         * 
         * ******************************/
        public void replaceVeteranInstances()
        {
            //  "Veteran-Owned,Veteran [oO]wned,veteran [oO]wned,veteran-[oO]wned"," Veteran-owned", "Veteran-owned should be capitalized and have a hyphen", False, True, False
            // "Veterans [iI]ntegrated [sS]ervice network,veterans [iI]ntegrated [sS]ervice [nN]etwork,Veterans integrated [sS]ervice Network,Veterans [iI]ntegrated service Network"), " Veterans Integrated Service Network", " Veterans Integrated Service Network should be capitalized", False, True, False


           // "veteran>,veterans>", "Veteran", "Veteran(s) should be capitalized", True, False, False
            apply_changes_to_word_permutations("Veteran-Owned,Veteran [oO]wned,veteran [oO]wned,veteran-[oO]wned", " Veteran-owned", "Veteran-owned should be capitalized and have a hyphen", "False, True, False");
            apply_changes_to_word_permutations("Veterans [iI]ntegrated [sS]ervice network,veterans [iI]ntegrated [sS]ervice [nN]etwork,Veterans integrated [sS]ervice Network,Veterans [iI]ntegrated service Network", " Veterans Integrated Service Network", " Veterans Integrated Service Network should be capitalized", "False, True, False");
            //apply_changes_to_word_permutations("veteran,veterans","Veteran","Veterans(s) should be capitalized","True,False,False");
            FindReplaceAndCommentWithWildCards("veteran", "Veteran" ,"Veterans(s) should be capitalized");
        

        }

        public void replaceFederalInstances()
        {
           
            apply_changes_to_word_permutations("Federal government, federal Government, federal government", " Federal Government", "Federal Government should be capitalized", "False, True, False");
            FindReplaceAndCommentWithWildCards("federal", "Federal", "Federal should be capitalized");

        }

        public void replaceCongressInstances()
        {
           // "members of congress,Members of congress,members of Congress"), " Members of Congress", "Members of Congress should be capitalized"

            apply_changes_to_word_permutations("members of congress,Members of congress,members of Congress", " Members of Congress", "Members of Congress should be capitalized", "False, True, False");
            FindReplaceAndCommentWithWildCards("congress", "Congress", "Congress / Congressional should be capitalized");

        }

        public void commentWebInstances()
        {
            //"Website,[wW]eb site"), "website(s) should not be capitalized nor written as two words"
            //"<Web[ -][bB]ased>,<web[ -]Based>"), "web-based should not be capitalized"
           // comment_changes_to_word_permutations("Website|[wW]eb site","Website,[wW]eb site","","website(s) should not be capitalized nor written as two words","true,false,true");


            //comment_changes_to_word_permutations("Web[ -][bB]ased|<web[ -]Based>|<web[ -]based>", "<Web[ -][bB]ased>,<web[ -]Based>,<web[ -]based>", "", "web-based should not be capitalized", "true,false,true");

            string[] textToFind = new string[] { "Website", "[wW]eb site" };
            foreach (var text in textToFind)
            {
                FindAndCommentWithWildCards(text, "website(s) should not be capitalized nor written as two words");
            }


            textToFind = new string[] { "<Web[ -][bB]ased>", "<web[ -]Based>", "<web[ -]based>" };
            foreach(var text in textToFind)
            {
                FindAndCommentWithWildCards(text, "web-based should not be capitalized");
            }

            //this needs come last because....
            
            FindAndCommentWithWholeWord("Web","web should not be capitalized");



        }

        public void commentCentralOffice()
        {
            //Central office| central[oO]ffice", "Central office, central[oO]ffice

            string[] textToFind = new string[] { "Central Office", "central[oO]ffice","Central office","central [oO]ffice" };
            foreach (var text in textToFind)
            {
                FindAndCommentWithWildCards(text, "Only capitalized if it's referring to the VA");
            }

            
        }

        public void commentCoWorkers()
        {
            
            //coworkers|Coworkers|Co workers|co-workers|co workers|Co-Workers|co-Workers|co Workers", "coworkers,Coworkers,Co workers,co-workers,co workers,Co-Workers,co-Workers,co Workers", 
            string[] textToFind = new string[] { "coworkers", "Coworkers", "Co workers", "co-workers","co workers", "Co-Workers","co-Workers"};
            foreach (var text in textToFind)
            {
                FindAndCommentWithWildCards(text, "Only capitalized if it's referring to the VA");
            }


        }
        public void commentGovernment()
        {

            //coworkers|Coworkers|Co workers|co-workers|co workers|Co-Workers|co-Workers|co Workers", "coworkers,Coworkers,Co workers,co-workers,co workers,Co-Workers,co-Workers,co Workers", 
            string[] textToFind = new string[] { "government", "Government"};
            foreach (var text in textToFind)
            {
                FindAndCommentWithWildCards(text, "Only capitalized if it's referring to the VA");
            }


        }







        /**************************
         * This method is used for rule finding
         * numbers under 10 in a sentence....
         * it's coupled with the method call
         * in the button click handler in
         * myUserControl.cs
         * 
         * ***********************/
        public void processSentences()
        {
            

            string[] wordsZeroThrough9 = new string[] {"zero","one","two","three","four","five","six","seven","eight","nine"};

            bool nineAndLower = false;
            bool tenAndHigher = false;
            bool spelledOutWordsUnder10 = false;

            Microsoft.Office.Interop.Word.Range wordRange = null;
            Microsoft.Office.Interop.Word.Range wordRange1 = null;
            Microsoft.Office.Interop.Word.Range wordRange2 = null;
            Range sentenceString = null;

            Word.Document document = this.Application.ActiveDocument;
            var sentenceCount = document.Sentences.Count;

            //Begin looping through each sentence in the document
            for (int i = 1; i <= sentenceCount; i++)
            {

                 nineAndLower = false;
                 tenAndHigher = false;
                 spelledOutWordsUnder10 = false;

                wordRange = this.Application.ActiveDocument.Sentences[i];

                //We need to save the sentence to here because the word range will change and it will
                //only comment one number, but we want to comment the entire sentence
                

                wordRange1 = this.Application.ActiveDocument.Sentences[i];
                wordRange2 = this.Application.ActiveDocument.Sentences[i];

                sentenceString = this.Application.ActiveDocument.Sentences[i];


                //MessageBox.Show(sentenceString);

                wordRange.Find.Text = "<[0-9]{1}>";

                wordRange.Find.Forward = true;
                wordRange.Find.Wrap = WdFindWrap.wdFindStop;

                //don't forget the Replace argument
                wordRange.Find.Execute(MatchWildcards: true);

                if (wordRange.Find.Found)
                {
                    nineAndLower = true;
                   // MessageBox.Show("9 and Lower Find");
                }

                wordRange1.Find.Text = "<[0-9]{2,10}>";

                //wordRange.Find.Replacement.Text = replacementText;//that's the right way to write
                wordRange1.Find.Forward = true;
                wordRange1.Find.Wrap = WdFindWrap.wdFindStop;

                //don't forget the Replace argument
                wordRange1.Find.Execute(MatchWildcards: true);

                if (wordRange1.Find.Found)
                {
                    tenAndHigher = true;
                   // MessageBox.Show("Ten and Higher True");
                }
                foreach(var word in wordsZeroThrough9)
                {
                    wordRange2.Find.Text = word;
                    wordRange2.Find.Forward = true;
                    wordRange2.Find.Wrap = WdFindWrap.wdFindStop;
                    wordRange2.Find.Execute(MatchCase:false);
                    if (wordRange2.Find.Found)
                    {
                        //MessageBox.Show("Found a spelled out word");
                        spelledOutWordsUnder10 = true;
                        break;
                    }
                }
              
                if (nineAndLower == true && tenAndHigher == false)
                {
                    object text = "Digits need to be spelled out - cme";
                    Word.Range rng = this.Application.ActiveDocument.Range(sentenceString.Start,sentenceString.End);
                    //rng.Text = ReplacementText;
                    document.Comments.Add(
                    rng, ref text);
                }
                if(spelledOutWordsUnder10 == true && tenAndHigher == true)
                {
                    object text = "Digits shouldn't be spelled out if there are numbers > 9 - cme";
                    //Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                    Word.Range rng = this.Application.ActiveDocument.Range(sentenceString.Start, sentenceString.End);
                    //rng.Text = ReplacementText;
                    document.Comments.Add(
                    rng, ref text);
                }

            }
        }
        public void FindAndReplaceSpacesAroundHyphens()
        {
            Microsoft.Office.Interop.Word.Range wordRange = null;
            Word.Document document = this.Application.ActiveDocument;
            wordRange = document.Content;
            wordRange.Find.ClearFormatting();
            wordRange.Find.ClearAllFuzzyOptions();
            wordRange.Find.Replacement.ClearFormatting();
            wordRange.Find.IgnoreSpace = true;
            wordRange.Find.MatchCase = false;
            //wordRange.Find.MatchWholeWord = optionValues[2];
            wordRange.Find.MatchWildcards = true;
            wordRange.Find.Text = " - ";
            wordRange.Find.Execute();
            //Regex reg = new Regex(TextToFind);
            while (wordRange.Find.Found)
            {
                //Match matchedItem = reg.Match(wordRange.Text);
                object text = "No Hypens";
                //if (matchedItem.Success)
                //{
                //wordRange.Text = ReplacementText;
                Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                rng.Text = "-";
                document.Comments.Add(
                rng, ref text);
                wordRange.Find.ClearFormatting();
                //}


                // Next Find
                wordRange.Find.Execute(FindText: " - ", MatchCase: false, MatchWildcards: false);
            }
        }
       
        public void FindAndReplaceWildcardPlayGround(string wildCardText, string replacementText, string commentMessage)
        {
            Microsoft.Office.Interop.Word.Range wordRange = null;
            Word.Document document = this.Application.ActiveDocument;
            wordRange = document.Content;            
            wordRange.Find.ClearFormatting();
            wordRange.Find.ClearAllFuzzyOptions();
            wordRange.Find.Replacement.ClearFormatting();
            wordRange.Find.IgnoreSpace = true;
            wordRange.Find.MatchCase = false;
            wordRange.Find.MatchWildcards = true;
            wordRange.Find.Text = wildCardText;
            wordRange.Find.Execute();
            while (wordRange.Find.Found)
            {
                object commentText = commentMessage;                
                Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                rng.Text = replacementText;
                document.Comments.Add(
                rng, ref commentText);
                wordRange.Find.ClearFormatting();
                
                // Next Find
                wordRange.Find.Execute(FindText: wildCardText, MatchCase: false, MatchWildcards: true);
            }
        }

        public List<Tuple<int, string, string, string, string, string>> ProcessDocument()
        {
            Word.Document document = this.Application.ActiveDocument;
            string textToSearch = document.Content.Text;
            // Item1 = method to use, Item2 = Regex, Item3 = Find search item, Item4 = replacement, Item5 = comments, Item6 = search settings, e.g. whole word          
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
        {
            //The contents of the returned tuple e.g. Item 2 = regex, item1 is the method
            //Item1  = method to use, Item2 = Regex, Item3 = Find search item, Item4 = replacement, Item5 = comments, Item6 = search settings, e.g. MatchWholeCase
            //Method 1 - apply_changes_to_word_permutations //This doesn't exist, I must create according to VBA Code
            //Method 2 - comment_on_change_to_word_permutations
            //Method 3 - replace_with_comments
            //Method 4 - add_comments
            //Method 5 - phone number replace

            //This is Marjories Code, it returns a list of tuples with all the things we need to search for
            List<Tuple<int, string, string, string, string, string>> styleArray = new List<Tuple<int, string, string, string, string, string>>();

           // styleArray.Add(new Tuple<int, string, string, string, string, string>(3, "veteran", "veteran", "Veteran", "Veteran(s) should be capitalized", "true, false, false"));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "Department-Wide|[dD]epartment [wW]ide|department-[wW]ide", "Department-Wide,[dD]epartment [wW]ide,department-[wW]ide", " Department-wide", "Department-wide should be capitalized and have a hyphen", "False, True, False"));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(5, @"\bnation\b", "nation", "Nation", "Nation should be capitalized", "True, False, True"));
            //styleArray.Add(new Tuple<int, string, string, string, string, string>(3, "congress", "congress", "Congress", "Congress / Congressional should be capitalized", "True, False, False"));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "[0-9]{10}|[(][1-9]{3}[)][1-9]{3}-[1-9]{4}|[(][1-9]{3}[)]-[1-9]{3}-[1-9]{4}|[1-9]{3}.[1-9]{3}.[1-9]{4}", "[0-9]{10},[(][1-9]{3}[)][1-9]{3}-[1-9]{4},[(][1-9]{3}[)]-[1-9]{3}-[1-9]{4},[1-9]{3}.[1-9]{3}.[1-9]{4}", null, "phone number should be in the format XXX-XXX-XXXX", null));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "service member|[sS]ervice Member", "service member,[sS]ervice Member", " Service member", "Service member(s) should be capitalized", "False, True, False"));
            //styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "members of congress|Members of congress|members of Congress", "members of congress,Members of congress,members of Congress", " Members of Congress", "Members of Congress should be capitalized", "True, False, True"));
            //styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "coworkers|Coworkers|Co workers|co-workers|co workers|Co-Workers|co-Workers|co Workers", "coworkers,Coworkers,Co workers,co-workers,co workers,Co-Workers,co-Workers,co Workers", " Co-workers", " Co-workers should be capitalized", "True, False, True"));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "Internet", "Internet", null, "internet should not be capitalized", null));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "Fiscal [yY]ear|[fF]iscal Year", "Fiscal [yY]ear,[fF]iscal Year", "fiscal year", "fiscal year should not be capitalized.", "True, False,False"));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "Intranet", "Intranet", null, "intranet should not be capitalized", null));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "[A-Z,a-z]{1,15} years old", "[A-Z,a-z]{1,15} years old", null, "Ages should be preceded by a digit.", null));
            // styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "Web", "<Web>", null, "web should not be capitalized", null));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "Armed forces|armed [fF]orces", "armed [fF]orces,Armed forces", " Armed Forces", "Armed Forces should be capitalized", "False, True, False"));
            //styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "[eE]-mail|Email", "[Ee]-mail,Email", null, "email should not be capitalized nor have a hyphen", null));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(1, @"Alzheimer's Disease|alzheimer's disease|alzheimer's Disease|Alzheimers Disease|Alzheimers disease|alzheimers disease|alzheimers Disease", @"Alzheimer's Disease,alzheimer's disease,alzheimer's Disease,Alzheimers Disease,Alzheimers disease,alzheimers disease,alzheimers Disease", @" Alzheimer's disease", "Alzheimer's should be capitalized", "True, False, True"));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "governmentWide|governmentwide|government-wide|Government-wide|Government-Wide", "governmentWide,governmentwide,government-wide,Government-wide,Government-Wide", " Governmentwide", "Governmentwide should be capitalized", "False, True, False"));
            //styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "Veterans integrated service network|veterans integrated service network", "Veterans integrated service network", " Veterans Integrated Service Network", " Veterans Integrated Service Network should be capitalized", "True, False, True"));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "Hepatitis [cC]|hepatitis c|hepatitis-c|hepatitis-C", "Hepatitis [cC],hepatitis c,hepatitis-c,hepatitis-C", null, "hepatitis C should not be capitalized", null));
           // styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "Website|[wW]eb site", "Website,[wW]eb site", null, "website(s) should not be capitalized nor written as two words", null));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "health Care|Health [cC]are|[hH]ealthcare", "health Care,Health [cC]are,[hH]ealthcare", "health care", "health care should not be capitalized", "False, True, False"));
            //styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "Web[ -][bB]ased|web[ -]Based", "<Web[ -][bB]ased>,<web[ -]Based>", null, "web-based should not be capitalized", null));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "Posttraumatic [sS]tress [dD]isorder|posttraumatic Stress [dD]isorder|posttraumatic [sS]tress Disorder|Posttraumatic-[sS]tress-[dD]isorder|posttraumatic-Stress-[dD]isorder|posttraumatic-[sS]tress-Disorder", "Posttraumatic [sS]tress [dD]isorder,posttraumatic Stress [dD]isorder,posttraumatic [sS]tress Disorder,Posttraumatic-[sS]tress-[dD]isorder,posttraumatic-Stress-[dD]isorder,posttraumatic-[sS]tress-Disorder", null, "posttraumatic stress disorder should not be capitalized", null));
            // styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "Central office|central [oO]ffice", "Central office, central [oO]ffice", " Central Office", "Central Office should be capitalized", "False, True, False"));
            //styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "Central office|central [oO]ffice", "Central office, central [oO]ffice", null, "Only capitalized if it's referring to the VA", null));
            //styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "federal", "federal", "Federal", "Federal should be capitalized", "True, False, True"));
            //styleArray.Add(new Tuple<int, string, string, string, string, string>(1, @"\bfederal\b(?!\W+[gG]overnment\b)", "Federal government", "Federal Government", "Federal Government should be capitalized", "True, False, True"));
            //styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "Veteran-Owned|[vV]eteran [oO]wned|veteran-[oO]wned", "Veteran-Owned,[vV]eteran [oO]wned,veteran-[oO]wned", " Veteran-owned", "Veteran-owned should be capitalized and have a hyphen", "True, False, True"));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, @"[eE]xecutive [oO]rder", "[eE]xecutive [oO]rder", null, "Executive Order should be capitalized when using EO number", null));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, @"[mM]edical [cC]enter", "[mM]edical [cC]enter", null, "Medical center should be capitalized if preceded by the formal name of the facility", null));
            styleArray.Add(new Tuple<int, string, string, string, string, string>(2, "cosigners", "cosigners", null, "Should be other signatories", null));

            //Format time
            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "[0-9]{1,2}:[0-9]{2}[AamM.]{2,4}[ -]{1,3}[0-9]{1,2}:[0-9]{2}[AamM.]{2,4}", "[0-9]{1,2}:[0-9]{2}[AamM.]{2,4}[ -]{1,3}[0-9]{1,2}:[0-9]{2}[AamM.]{2,4}", null, "if time range are both in am, time should be written as X:XX-X:XX a.m. (e.g. 10:15-11:30 a.m. or 8-9 a.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "[0-9]{1,2}:[0-9]{2}[PpmM.]{2,4}[ -]{1,3}[0-9]{1,2}:[0-9]{2}[PpmM.]{2,4}", "[0-9]{1,2}:[0-9]{2}[PpmM.]{2,4}[ -]{1,3}[0-9]{1,2}:[0-9]{2}[PpmM.]{2,4}", null, "if time range are both in pm, time should be written as X:XX-X:XX p.m. (e.g. 1:15-2:30 p.m. or 4-6 p.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, " [0-9]{1,2} [AamM.]{2,4}[ -]{1,3}[0-9]{1,2} [AamM.]{2,4}", " [0-9]{1,2} [AamM.]{2,4}[ -]{1,3}[0-9]{1,2} [AamM.]{2,4}", null, "if time range are both in am, time should be written as X-X a.m. (e.g. 10:15-11:30 a.m. or 8-9 a.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, " [0-9]{1,2} [PpmM.]{2,4}[ -]{1,3}[0-9]{1,2} [PpmM.]{2,4}", " [0-9]{1,2} [PpmM.]{2,4}[ -]{1,3}[0-9]{1,2} [PpmM.]{2,4}", null, "if time range are both in pm, time should be written as X-X p.m. (e.g. 1:15-2:30 p.m. or 4-6 p.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, " [0-9]{1,2}[AamM.]{2,4}[ -]{1,3}[0-9]{1,2}[AamM.]{2,4}", " [0-9]{1,2}[AamM.]{2,4}[ -]{1,3}[0-9]{1,2}[AamM.]{2,4}", null, "if time range are both in am, time should be written as X-X a.m. (e.g. 10:15-11:30 a.m. or 8-9 a.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, " [0-9]{1,2}[PpmM.]{2,4}[ -]{1,3}[0-9]{1,2}[PpmM.]{2,4}", " [0-9]{1,2}[PpmM.]{2,4}[ -]{1,3}[0-9]{1,2}[PpmM.]{2,4}", null, "if time range are both in pm, time should be written as X-X p.m. (e.g. 1:15-2:30 p.m. or 4-6 p.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "12 [aA].[mM].|12 [aA][mM]|12 [aA].[mM]|12:00 [aA].[mM].|12:00 [aA][mM]|12:00 [aA].[mM]", "12 [aA].[mM].,12 [aA][mM],12 [aA].[mM],12:00 [aA].[mM].,12:00 [aA][mM],12:00 [aA].[mM]", " midnight", "midnight should be used instead of 12 a.m.", "False, True, False"));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(1, "12 [pP].[mM].|12 [pP][mM]|12 [pP].[mM]|12:00 [pP].[mM].|12:00 [pP][mM]|12:00 [pP].[mM]", "12 [pP].[mM].,12 [pP][mM],12 [pP].[mM],12:00 [pP].[mM].,12:00 [pP][mM],12:00 [pP].[mM]", " noon", "midnight should be used instead of 12 a.m.", "False, True, False"));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "[0-9]{1,2} [pPaA][mM]", "[0-9]{1,2} [pPaA][mM]", null, "time should use lowercase a.m./p.m. with periods in between (e.g. 8 a.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "[0-9]{1,2} [pPaA][mM]", "[0-9]{1,2} [pPaA][mM]", null, "time should use lowercase a.m./p.m. with periods in between (e.g. 8 a.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "[0-9]{1,2}[pPaA][mM]", "[0-9]{1,2}[pPaA][mM]", null, "time should be in format: XX:XX a.m./p.m. with a space between the digit and a.m./p.m. suffix (e.g. 8 a.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "[0-9]{1,2}[PpaA].[Mm].", "[0-9]{1,2}[PpaA].[Mm].", null, "time should be in format: XX:XX a.m./p.m. with a space between the digit and a.m./p.m. suffix(e.g. 8 a.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, " 0[0-9]", " 0[0-9]", null, "time should be written without a preceding zero (e.g. 7:15 a.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(3, ":00", ":00", " ", "time written without minutes should be written as hours only (e.g. 11 a.m.)", "False, False, True"));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "[0-9,.] - [0-9]", "[0-9,.] - [0-9]", null, "a time range should use a hyphen without surrounding spaces (e.g. 8-9 a.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "[0-9,.]- [0-9]", "[0-9,.]- [0-9]", null, "a time range should use a hyphen without surrounding spaces (e.g. 8-9 a.m.)", null));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(4, "[0-9,.] -[0-9]", "[0-9,.] -[0-9]", null, "a time range should use a hyphen without surrounding spaces (e.g. 8-9 a.m.)", null));

            //Conjuntions
            styleArray.Add(new Tuple<int, string, string, string, string, string>(3, ", and", ", and", " and", "When using commas to separate elements of a series, do not put a comma before the conjunction.'", "True"));

            styleArray.Add(new Tuple<int, string, string, string, string, string>(3, "  ", "  ", " ", "There were two spaces here and it's now one space.", "True"));

            return styleArray;
        }

        public void ReplaceWithCommentsWholeWord(string TextToFind, string ReplacementText, string CommentText, string settings)
        {

            //var found = false;
            List<bool> optionValues = new List<bool>();
            var functionsettings = settings.Split(',');
            foreach (var x in functionsettings)
            {
                optionValues.Add(Convert.ToBoolean(x));
            }
            Microsoft.Office.Interop.Word.Range wordRange = null;
            Word.Document document = this.Application.ActiveDocument;

            wordRange = document.Content;
            wordRange.Find.ClearFormatting();
            wordRange.Find.ClearAllFuzzyOptions();
            wordRange.Find.Replacement.ClearFormatting();
            wordRange.Find.IgnoreSpace = true;
            wordRange.Find.MatchCase = false;
            wordRange.Find.MatchWholeWord = true;
            wordRange.Find.MatchWildcards = false;
            wordRange.Find.Text = TextToFind;
            wordRange.Find.Execute();
            Regex reg = new Regex(TextToFind);
            while (wordRange.Find.Found)
            {
                Match matchedItem = reg.Match(wordRange.Text);
                object text = "[REPLACED]: " + CommentText + " -CME";
                if (matchedItem.Success)
                {
                    //wordRange.Text = ReplacementText;
                    Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                    rng.Text = ReplacementText;
                    document.Comments.Add(
                    rng, ref text);
                    wordRange.Find.ClearFormatting();
                }


                // Next Find
                wordRange.Find.Execute(FindText: TextToFind, MatchCase: false, MatchWildcards: false, MatchWholeWord:true);
            }

        }




        public void ReplaceWithComments(string TextToFind, string ReplacementText, string CommentText, string settings)
        {

            //var found = false;
            List<bool> optionValues = new List<bool>();
            var functionsettings = settings.Split(',');
            foreach (var x in functionsettings)
            {
                optionValues.Add(Convert.ToBoolean(x));
            }
            Microsoft.Office.Interop.Word.Range wordRange = null;
            Word.Document document = this.Application.ActiveDocument;            

            wordRange = document.Content;
            wordRange.Find.ClearFormatting();
            wordRange.Find.ClearAllFuzzyOptions();
            wordRange.Find.Replacement.ClearFormatting();
            wordRange.Find.IgnoreSpace = true;
            wordRange.Find.MatchCase = false;
            //wordRange.Find.MatchWholeWord = optionValues[2];
            wordRange.Find.MatchWildcards = true;
            wordRange.Find.Text = TextToFind;
            wordRange.Find.Execute();
            Regex reg = new Regex(TextToFind);
            while (wordRange.Find.Found)
            {
                Match matchedItem = reg.Match(wordRange.Text);
                object text = "[REPLACED]: " + CommentText + " -CME";
                if (matchedItem.Success)
                {
                    //wordRange.Text = ReplacementText;
                    Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                    rng.Text = ReplacementText;
                    document.Comments.Add(
                    rng, ref text);
                    wordRange.Find.ClearFormatting();
                }


                // Next Find
                wordRange.Find.Execute(FindText: TextToFind, MatchCase: false, MatchWildcards: true);
            }

        }

        public void ReplaceWithCommentsNonStyleArray(string TextToFind, string ReplacementText, string CommentText)
        {
            Microsoft.Office.Interop.Word.Range wordRange = null;
            Word.Document document = this.Application.ActiveDocument;
            wordRange = document.Content;
            wordRange.Find.ClearFormatting();
            wordRange.Find.ClearAllFuzzyOptions();
            wordRange.Find.Replacement.ClearFormatting();
            wordRange.Find.IgnoreSpace = true;
            wordRange.Find.MatchCase = false;
            wordRange.Find.MatchWildcards = true;
            wordRange.Find.Text = TextToFind;
            wordRange.Find.Execute();
            while (wordRange.Find.Found)
            {
                object text = "[REPLACED]: " + CommentText + " -CME";
                Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                rng.Text = ReplacementText;
                document.Comments.Add(
                rng, ref text);
                wordRange.Find.ClearFormatting();
               
                // Next Find
                wordRange.Find.Execute(FindText: TextToFind, MatchCase:false, MatchWildcards: true);
            }

        }

        public void ReplaceWithCommentsLoopThroughSentences(string TextToFind, string ReplacementText, string CommentText)
        {
            Microsoft.Office.Interop.Word.Range wordRange = null;
            Word.Document document = this.Application.ActiveDocument;
            var sentenceCount = document.Sentences.Count;
            for(int i = 1; i <= sentenceCount; i++)
            {
                wordRange = this.Application.ActiveDocument.Sentences[i];
               // MessageBox.Show(wordRange.Text);
            }
           
            wordRange = document.Content;
            wordRange.Find.ClearFormatting();
            wordRange.Find.ClearAllFuzzyOptions();
            wordRange.Find.Replacement.ClearFormatting();
            wordRange.Find.IgnoreSpace = true;
            wordRange.Find.MatchCase = false;
            wordRange.Find.MatchWildcards = true;
            wordRange.Find.Text = TextToFind;
            wordRange.Find.Execute();
            while (wordRange.Find.Found)
            {
                object text = CommentText;
                Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                rng.Text = ReplacementText;
                document.Comments.Add(
                rng, ref text);
                wordRange.Find.ClearFormatting();

                // Next Find
                wordRange.Find.Execute(FindText: TextToFind, MatchCase: false, MatchWildcards: true);
            }

        }



        public void AddComments(string regexTextToFind,string TextToFind, string ReplacementText, string CommentText, string settings)
        {
            //var found = false;

            Microsoft.Office.Interop.Word.Range wordRange = null;
            Word.Document document = this.Application.ActiveDocument;
            wordRange = document.Content;
            wordRange.Find.ClearFormatting();
            wordRange.Find.ClearAllFuzzyOptions();
            wordRange.Find.Replacement.ClearFormatting();
            wordRange.Find.IgnoreSpace = true;
            wordRange.Find.MatchCase = false;
            //wordRange.Find.MatchWholeWord = optionValues[2];
            wordRange.Find.MatchWildcards = true;
            wordRange.Find.Text = TextToFind;
            wordRange.Find.Execute();
            Regex reg = new Regex(regexTextToFind);
            while (wordRange.Find.Found)
            {
                Match matchedItem = reg.Match(wordRange.Text);
                object text = CommentText + " -CME";
                if (matchedItem.Success)
                {
                    //wordRange.Text = ReplacementText;
                    Word.Range rng = this.Application.ActiveDocument.Range(wordRange.Start, wordRange.End);
                    //rng.Text = ReplacementText;
                    document.Comments.Add(
                    rng, ref text);
                    wordRange.Find.ClearFormatting();
                }

                // Next Find
                wordRange.Find.Execute(FindText: TextToFind, MatchCase: false, MatchWildcards: true);
            }

        }

        /*************************************
         * cannot use matchcase and wildcards together
         * 
         * 
         * ***************************************/
        public void CommentWithoutReplace(string WordToComment, string message)
        {
            Word.Document document = this.Application.ActiveDocument;
            Word.Range rng = document.Content;

            rng.Find.ClearFormatting();
            rng.Find.Forward = true;
            rng.Find.MatchWildcards = true;  //This was just added 5/3/2022, may need to remove
            rng.Find.Text = WordToComment;
            

            rng.Find.Execute(
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);

            while (rng.Find.Found)
            {
                rng.Find.MatchWildcards = true;   //This was just added, may need to delete 5/3/2023
                
                rng.Find.Execute(
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);

                //object text = WordToComment + " " + message + " -CME";
                object text = message + " -CME";
                object start = rng.Start;
                object end = rng.End;

                this.Application.ActiveDocument.Comments.Add(
                    Application.ActiveDocument.Range(rng.Start, rng.End), ref text);
            }

        }

        
        public void CommentWithoutReplaceCheckSentence(string WordToComment, string message)
        {
           
            Word.Document document = this.Application.ActiveDocument;
            Word.Range rng = document.Content;

            var sentenceCount = document.Sentences.Count;
            for (int i = 1; i <= sentenceCount; i++)
            {
                rng = this.Application.ActiveDocument.Sentences[i];
                rng.Find.ClearFormatting();
                rng.Find.Forward = true;
                rng.Find.MatchWildcards = true;  //This was just added 5/3/2022, may need to remove
                rng.Find.Text = WordToComment;

                rng.Find.Execute(
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);
                if (rng.Find.Found)
                {
                    numberCounter++;
                }
                if (numberCounter >=3)
                {
                    while (rng.Find.Found)
                    {
                        rng.Find.MatchWildcards = true;   //This was just added, may need to delete 5/3/2023
                        rng.Find.Execute(
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing);

                        //object text = WordToComment + " " + message + " -CME";
                        object text = message + " -CME";
                        object start = rng.Start;
                        object end = rng.End;

                        this.Application.ActiveDocument.Comments.Add(
                            Application.ActiveDocument.Range(rng.Start, rng.End), ref text);
                    }
                }                

            }
            
           // MessageBox.Show(numberCounter.ToString());
            numberCounter = 0;


        }

        public void apply_changes_to_word_permutations(string TextToFind, string ReplacementText, string CommentText, string settings)
        {
            var textToFindArray = TextToFind.Split(',');
            foreach (var text in textToFindArray)
            {
                ReplaceWithComments(text, ReplacementText, CommentText, settings);
            }
        }

        public void comment_changes_to_word_permutations(string regexTextToFind, string TextToFind, string ReplacementText, string CommentText, string settings)
        {
            var textToFindArray = TextToFind.Split(',');
            foreach (var text in textToFindArray)
            {
                AddComments(regexTextToFind, text, ReplacementText, CommentText, settings);
            }
        }

        public void FormatDate()
        {
            string[] monthsArray = new string[] { "[Jj]anuary", "[Ff]ebruary", "[Mm]arch", "[Aa]pril", "[Mm]ay", "[Jj]une", "[Jj]uly",
            "[Aa]ugust", "[Ss]eptember", "[Oo]ctober","[Nn]ovember", "[Dd]ecember"};

            foreach (var x in monthsArray)
            {
                CommentWithoutReplace("[A-Z,a-z][th, st, nd, rd] of " + x, "the date should be written as DD(st/nd/rd/th) of " + x + " (e.g. 11th of November)");
                CommentWithoutReplace("[A-Z,a-z][th, st, nd, rd] " + x, "the date should be written as DD(st/nd/rd/th) of " + x + " (e.g. 11th of November)");
                CommentWithoutReplace(x + " [0-9]{1,2}[A-Za-z]{2}", "the date should be written as " + x + " DD (e.g May 1)");

            }


        }
        public void FormatNumbersUnder10()
        {
            string[] numbersArray = new string[] { "zero","one","two","three",
            "four", "five", "six","seven", "eight","nine"};

            int[] digitsArray = new int[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };


            foreach (var x in digitsArray)
            {
                //CommentWithoutReplace(" " + x + " ", "numbers under ten should be written out in words (when describing amounts of objects e.g. nine Veterans)");
                //CommentWithoutReplace(x.ToString(), "numbers under ten should be written out in words (when describing amounts of objects e.g. nine Veterans)");
                CommentWithoutReplaceCheckSentence(x.ToString(), "Numbers under ten...");
            }


        }

        public void DollarSymbolFollowedByDigits()
        {
            string[] numbers = new string[] { "zero","one","two","three",
            "four", "five", "six","seven", "eight","nine"};
            int[] digitsArray = new int[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            for (int i = 0; i < numbers.Length; i++)
            {

                ReplaceWithCommentsNonStyleArray("$" + numbers[i], "$" + digitsArray[i], "$ signs should be written as digits");
                //replace_with_comments num(i) + " dollars", "$" + digit(i), " dollar amounts should be written as digits", False
                //ReplaceWithCommentsNonStyleArray(numbers[i] + " dollars", "$" + digitsArray[i], "$ signs should be written as digits");
                ReplaceWithCommentsNonStyleArray("[0 - 9]{1,15} dollars", "$" + digitsArray[i], "$ signs should be written as digits");

            }

        }
        //comment_symbol_should_be_preceeded_by_digits
        public void SymbolShouldBePreceededByDigitis()
        {

            string[] numericSymbolsPreceededByNumber = { " percent", " %", " cent", " degrees Fahrenheit", " degrees Celsius", "°F", "°C", "-\"", "," };
            string[] numbers = { "zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine" };
            //SymbolsPreecedByNumber = Split("%,%,¢, years old,°F,°C,°F,°C,-", ",")
            string[] symbolsPreceededByNumber = { "%", "%", "¢", "°F", "°C", "°F", "°C", "-\"", "," };
            string[] numbersDigits = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };


            for (int i = 0; i < numericSymbolsPreceededByNumber.Length; i++)
            {
                for (int j = 0; j < numbers.Length; j++)
                {
                    ReplaceWithCommentsNonStyleArray(numbers[j] + numericSymbolsPreceededByNumber[i], numbersDigits[i] + symbolsPreceededByNumber[i],
                        numericSymbolsPreceededByNumber[i] + " should be preceded by a digit");
                }
            }
        }

        public void DeleteAllComments(bool showMessageBox)
        {
            if (Application.ActiveDocument.Comments.Count != 0)
            {
                this.Application.ActiveDocument.DeleteAllComments();
                if (showMessageBox)
                {
                   // ShowAllComments();
                    MessageBox.Show("All Comments Have Been Cleared");
                }
            }
            else
            {
                if (showMessageBox)
                {
                    MessageBox.Show("There are No Comments to Delete");
                }

            }

        }
        /*public void ShowAllComments()
        {
            foreach (var comment in this.Application.ActiveDocument.Comments)
            {
                MessageBox.Show();

            }

        }*/



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
