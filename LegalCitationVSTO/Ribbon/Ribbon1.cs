﻿using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace LegalCitationVSTO
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            bool isEnabled = GlobalState.Enabled;
            GlobalState.Enabled = !isEnabled;

            Dictionary<bool, string> textDict = new Dictionary<bool, string>()
            {
                { false, "Enable" },
                { true, "Disable" }
            };

            Globals.Ribbons.Ribbon1.toggleButton.Label = textDict[isEnabled];
            e.Control.Context();
        }
    }
}
