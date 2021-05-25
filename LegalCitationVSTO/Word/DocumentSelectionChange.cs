using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using LegalCitationVSTO.Service.StringService;
using Microsoft.Office.Interop.Word;
using Tools = Microsoft.Office.Tools.Word;

namespace LegalCitationVSTO
{
    /// <summary>
    /// When the user copies text from the web extension.
    /// </summary>
    public partial class ThisAddIn
    {
        private void DocumentSelectionChange(Document nativeDocument)
        {
            Tools.Document vstoDoc = Globals.Factory.GetVstoObject(nativeDocument);
            vstoDoc.SelectionChange += new Tools.SelectionEventHandler(this.ThisDocument_SelectionChange);
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

            string footnoteText = this.stringService.ExtractFootnoteText(text);
            if (footnoteText == null) return;

            // Remove footnote tokens from paragraph
            string footnoteTextWithToken = this.stringService.FindMatch(text, StringService.FootnoteRegex);
            this.SearchReplaceFootnoteFromParagraph(paragraph, footnoteTextWithToken);

            Footnotes footnotes = this.Application.Selection.Footnotes;
            Footnote footnote = footnotes.Add(Range: paragraph.Range, Text: footnoteText);
            footnote.Range.Font.Color = WdColor.wdColorRed;
        }

        private void SearchReplaceFootnoteFromParagraph(Paragraph paragraph, string footnoteTextWithToken)
        {
            // MessageBox.Show($"{paragraph.Range.Text}");
            Find findObject = paragraph.Range.Find;
            findObject.ClearFormatting();

            bool found = findObject.Execute(
                Replace: WdReplace.wdReplaceAll,
                FindText: footnoteTextWithToken,
                ReplaceWith: string.Empty);

            Console.WriteLine(found);
        }
    }
}
