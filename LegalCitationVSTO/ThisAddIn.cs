using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LegalCitationVSTO.Service.StringService;

namespace LegalCitationVSTO
{
    public partial class ThisAddIn
    {
        // Instantiate services
        readonly IString stringService = new StringService();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(DocumentSelectionChange);
            ((ApplicationEvents4_Event)Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(DocumentSelectionChange);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
