using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LegalCitationVSTO.Service.StringService;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LegalCitationVSTO.Test
{
    [TestClass]
    public class StringServiceTest
    {
        [TestMethod]
        public void ExtractFootnoteText()
        {
            IString stringService = new StringService();
            const string source = "Personal Data Protection Act 2012 (No. 26 of 2012) s 14(1)";
            string copiedText = $"An individual has not given consent under this Act for the collection__FOOTNOTE__{source}__/FOOTNOTE__";
            string footnoteText = stringService.ExtractFootnoteText(copiedText);
            Assert.AreEqual(footnoteText, source);
        }
    }
}
