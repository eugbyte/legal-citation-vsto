using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using WK.Libraries.SharpClipboardNS;
using static WK.Libraries.SharpClipboardNS.SharpClipboard;

namespace LegalCitationVSTO
{
    /// <summary>
    /// When the user copies text to windows clipboard.
    /// </summary>
    public partial class ThisAddIn
    {
        private readonly SharpClipboard clipboard = new SharpClipboard();

        private void ClipboardChange(Document document)
        {
            Debug.WriteLine("Detection started ...");
            this.clipboard.ClipboardChanged += this.OnClipboardChange;
        }

        /// <summary>
        /// When the user copies text from pdf files.
        /// </summary>
        private void OnClipboardChange(object sender, ClipboardChangedEventArgs e)
        {
            Debug.WriteLine("Copy detected");

            if (e.ContentType != ContentTypes.Text) return;

            string copiedText = this.clipboard.ClipboardText;
            string applicationName = e.SourceApplication.Name;
            string title = e.SourceApplication.Title;
            int id = e.SourceApplication.ID;

            Debug.WriteLine(copiedText);
            Debug.WriteLine(id);
            Debug.WriteLine(applicationName);
            Debug.WriteLine(title);

            if (!title.Contains(".pdf")) return;

            Document doc = this.Application.ActiveDocument;
            RibbonButton button = Globals.Ribbons.Ribbon1.pdfButton;
            button.Label = title;
        }
    }
}
