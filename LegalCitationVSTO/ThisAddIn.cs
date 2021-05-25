using System;
using System.Diagnostics;
using LegalCitationVSTO.Service.StringService;
using Microsoft.Office.Interop.Word;
using WK.Libraries.SharpClipboardNS;

namespace LegalCitationVSTO
{
    /// <summary>
    /// Entry point.
    /// </summary>
    public partial class ThisAddIn
    {
        // Instantiate services
        private readonly IString stringService = new StringService();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // This beta software will be availabe for one month
            if (DateTime.Now > new DateTime(day: 28, month: 6, year: 2021))
            {
                return;
            }

            this.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(this.DocumentSelectionChange);
            ((ApplicationEvents4_Event)this.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(this.DocumentSelectionChange);

            this.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(this.ClipboardChange);
            ((ApplicationEvents4_Event)this.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(this.ClipboardChange);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(this.ThisAddIn_Startup);
            this.Shutdown += new EventHandler(this.ThisAddIn_Shutdown);
        }

        #endregion
    }
}
