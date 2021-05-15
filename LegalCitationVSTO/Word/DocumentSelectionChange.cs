using LegalCitationVSTO.Service.StringService;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Tools = Microsoft.Office.Tools.Word;

namespace LegalCitationVSTO
{
    public partial class ThisAddIn
    {
        readonly IString stringService = new StringService();
        private void DocumentSelectionChange(Document nativeDocument)
        {
            Tools.Document vstoDoc = Globals.Factory.GetVstoObject(nativeDocument);
            vstoDoc.SelectionChange += new Tools.SelectionEventHandler(ThisDocument_SelectionChange);           
        }

        private void ThisDocument_SelectionChange(object sender, Tools.SelectionEventArgs e)
        {
            if (!GlobalState.Enabled) return;

            Range previous = e.Selection.Previous();

            if (previous == null || previous.Paragraphs == null) return;
            Paragraph paragraph = previous.Paragraphs[1];
            string text = paragraph.Range.Text;

            // Selection Event is expensive
            // Return immediately if no citation match
            int matches = Regex.Matches(text, StringService.FootnoteRegex).Count;
            if (matches != 1) return;

            string footnoteText = stringService.ExtractFootnoteText(text);
            if (footnoteText == null) return;

            // Remove footnote tokens from paragraph
            paragraph.Range.Text = Regex.Replace(text, StringService.FootnoteRegex, replacement: "");

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
