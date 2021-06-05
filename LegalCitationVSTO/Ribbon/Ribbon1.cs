using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace LegalCitationVSTO
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            RibbonButton button = Globals.Ribbons.Ribbon1.toggleButton;
            button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            bool isEnabled = GlobalState.Enabled;
            GlobalState.Enabled = !isEnabled;

            Dictionary<bool, string> textDict = new Dictionary<bool, string>()
            {
                { false, "Enable" },
                { true, "Disable" },
            };

            RibbonButton button = Globals.Ribbons.Ribbon1.toggleButton;

            button.Label = textDict[isEnabled];
            e.Control.Context();
        }

        private void PdfButton_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton button = Globals.Ribbons.Ribbon1.pdfButton;
            string footnoteText = button.Label;
            Footnotes footnotes = Globals.ThisAddIn.Application.Selection.Footnotes;
            Footnote footnote = footnotes.Add(Text: footnoteText);
            footnote.Range.Font.Color = WdColor.wdColorRed;

            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
            
        }
    }
}
