using LegalCitationVSTO.Service.StringService;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
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
            if (!Regex.IsMatch(text, StringService.FootnoteRegex)) return;

            string footnoteText = stringService.ExtractFootnoteText(text);
            if (footnoteText == null) return;

            // Remove footnote tokens from paragraph
            // This method of replacing is removing previous footnotes and styling
            // Need to zone in on the match
            // paragraph.Range.Text = Regex.Replace(text, StringService.FootnoteRegex, replacement: "");
            string footnoteTextWithToken = stringService.FindMatch(text, StringService.FootnoteRegex);
            SearchReplaceFootnoteToken(paragraph, footnoteTextWithToken);

            Application.Selection.Footnotes.Add(Range: paragraph.Range, Text: footnoteText);

            return;
        }

        private void SearchReplaceFootnoteToken(Paragraph paragraph, string footnoteText)
        {
            // MessageBox.Show($"{paragraph.Range.Text}");
            Find findObject = paragraph.Range.Find;
            findObject.ClearFormatting();

            bool found =  findObject.Execute(
                Replace: WdReplace.wdReplaceAll, 
                FindText: footnoteText
            );

            MessageBox.Show(found.ToString());
        }

    }
}
