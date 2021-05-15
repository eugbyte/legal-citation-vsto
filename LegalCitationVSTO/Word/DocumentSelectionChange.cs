using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Interop = Microsoft.Office.Interop.Word;
using Tools = Microsoft.Office.Tools.Word;

namespace LegalCitationVSTO
{
    public partial class ThisAddIn
    {
        private void DocumentSelectionChange(Interop.Document nativeDocument)
        {
            Tools.Document vstoDoc = Globals.Factory.GetVstoObject(nativeDocument);
            vstoDoc.SelectionChange += new Tools.SelectionEventHandler(ThisDocument_SelectionChange);
        }

        private void ThisDocument_SelectionChange(object sender, Tools.SelectionEventArgs e)
        {
            if (!GlobalState.Enabled) return;

            Document document = e.Selection.Document;
            Range previous = e.Selection.Previous();

            if (previous == null || previous.Paragraphs == null) return;

            if (e.Selection.Sentences?.Count == 0) return;

            int numSentences = e.Selection.Sentences.Count;
            int numFootnotes = e.Selection.Sentences[numSentences].Footnotes.Count;
            if (numFootnotes != 0) return;

            Paragraph paragraph = previous.Paragraphs[1];
            string text = paragraph.Range.Text;

            string pattern = @"__FOOTNOTE__.+__/FOOTNOTE__";
            MatchCollection mc = Regex.Matches(text, pattern);
            if (mc.Count != 1) return;

            // Extract citation without footnote tokens
            Match match = mc[0];
            string footnoteText = match.Value;
            footnoteText = Regex.Replace(footnoteText, "__FOOTNOTE__", "");
            footnoteText = Regex.Replace(footnoteText, "__/FOOTNOTE__", "");

            // Remove footnote tokens from paragraph
            paragraph.Range.Text = Regex.Replace(text, pattern, replacement: "");

            // When text is replaced, strangely a linebreak is inserted
            // Go back to the previous line before appending the footnote
            paragraph.Range.Previous().Select();
            Application.Selection.Select();
            Application.Selection.Text = "";
            Application.Selection.Footnotes.Add(Range: paragraph.Range, Text: footnoteText);

            // MessageBox.Show($"previous: {text}");
            return;
        }
    }
}
