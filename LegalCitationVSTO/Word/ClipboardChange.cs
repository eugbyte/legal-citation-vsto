using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using RestSharp;
using WK.Libraries.SharpClipboardNS;
using static WK.Libraries.SharpClipboardNS.SharpClipboard;
using Task = System.Threading.Tasks.Task;

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
            int id = title.GetHashCode();

            Debug.WriteLine(id);
            Debug.WriteLine(applicationName);
            Debug.WriteLine(title);
            Debug.WriteLine(copiedText);

            if (!title.Contains(".pdf")) return;

            RibbonButton button = Globals.Ribbons.Ribbon1.pdfButton;
            Dictionary<int, string> titleDict = GlobalState.TitleDict;

            if (!titleDict.ContainsKey(id))
            {
                string newTitle = Regex.Replace(title, @".pdf.+", string.Empty);
                button.Label = newTitle;

                // Note that this is async
                // So will be completed after code block below
                Task.Run(async () =>
                {
                    try
                    {
                        string fetchedTitle = await this.GetFullTitle();
                        Debug.WriteLine(fetchedTitle);
                    } catch (Exception error)
                    {
                        Debug.WriteLine(error.Message);
                    }

                    // When you get the actual API, shift the code below in the try block
                    titleDict.Add(id, newTitle);
                    button.Label = titleDict[id];
                });
            }
            else
            {
                button.Label = titleDict[id];
            }
        }

        /// <summary>
        /// Mock the retrieval of the actual title.
        /// </summary>
        private async Task<string> GetFullTitle()
        {
            Debug.WriteLine("starting fetch...");
            RestClient client = new RestClient("https://jsonplaceholder.typicode.com");
            RestRequest request = new RestRequest("todos/1", DataFormat.Json);
            string result = await client.GetAsync<string>(request);
            return result;
        }
    }
}
