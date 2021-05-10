using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Interop = Microsoft.Office.Interop.Word;
using Tools = Microsoft.Office.Tools.Word;

namespace LegalCitationVSTO
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(DocumentSelectionChange);
            ((Interop.ApplicationEvents4_Event)this.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(DocumentSelectionChange);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void DocumentSelectionChange(Interop.Document nativeDocument)
        {
            Tools.Document vstoDoc = Globals.Factory.GetVstoObject(nativeDocument);
            vstoDoc.SelectionChange += new Tools.SelectionEventHandler(ThisDocument_SelectionChange);
        }

        private void ThisDocument_SelectionChange(object sender, Tools.SelectionEventArgs e)
        {
            Document document = e.Selection.Document;            
            Range previous = e.Selection.Previous();

            if (previous == null || previous.Paragraphs == null) return;
            int numParas = previous.Paragraphs.Count;

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

            paragraph.Range.Previous().Select();

            Application.Selection.Select();
            Application.Selection.Text = "Hello";
            Application.Selection.Footnotes.Add(Range: paragraph.Range, Text: footnoteText);

            // When text is replaced, strangely a linebreak is inserted

            //paragraph
            //    .Range                
            //    .Footnotes.Add(Range: paragraph.Range, Text: footnoteText);

            // e.Selection.Range.InsertParagraphAfter();
            // e.Selection.Next().Select();
            // MessageBox.Show($"previous: {text}");
            return;
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
