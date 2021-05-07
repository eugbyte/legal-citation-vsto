﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Interop = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;

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

        void ThisDocument_SelectionChange(object sender, Tools.SelectionEventArgs e)
        {
            Document document = e.Selection.Document;
            Range previous = e.Selection.Previous();
            if (previous == null || previous.Paragraphs == null) return;

            Paragraph paragraph = previous.Paragraphs[1];

            string text = paragraph.Range.Text;

            MessageBox.Show($"previous: {text}");

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