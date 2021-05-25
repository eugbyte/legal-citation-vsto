using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WK.Libraries.SharpClipboardNS;
using static WK.Libraries.SharpClipboardNS.SharpClipboard;

namespace LegalCitationVSTO
{
    /// <summary>
    /// When the user copies text from pdf files.
    /// </summary>
    public partial class ThisAddIn
    {
        private readonly SharpClipboard clipboard = new SharpClipboard();

        private void ClipboardChange(Document document)
        {
            Debug.WriteLine("Detection started ...");
            this.clipboard.ClipboardChanged += this.OnClipboardChange;
        }

        private void OnClipboardChange(object sender, ClipboardChangedEventArgs e)
        {
            Document doc = this.Application.ActiveDocument;

            Debug.WriteLine("Copy detected");

            // Is the content copied of text type?
            if (e.ContentType == ContentTypes.Text)
            {
                // Get the cut/copied text.
                string copiedText = this.clipboard.ClipboardText;
                string applicationName = e.SourceApplication.Name;
                string title = e.SourceApplication.Title;

                Debug.WriteLine(copiedText);
                Debug.WriteLine(applicationName);
                Debug.WriteLine(title);
            }
        }
    }
}
