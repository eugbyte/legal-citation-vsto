using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace LegalCitationVSTO.Service.StringService
{
    class StringService: IString
    {
        public static readonly string FootnoteRegex = @"__FOOTNOTE__.+__/FOOTNOTE__";

        public string FindMatch (string text, string pattern)
        {
            MatchCollection mc = Regex.Matches(text, pattern);
            if (mc.Count != 1) return null;

            Match match = mc[0];
            return match.Value;
        }

        public string ExtractFootnoteText (string text)
        {
            string footnoteText = FindMatch(text, FootnoteRegex);
            if (footnoteText == null) return  null;
            footnoteText = Regex.Replace(footnoteText, "__FOOTNOTE__", "");
            footnoteText = Regex.Replace(footnoteText, "__/FOOTNOTE__", "");
            return footnoteText;
        }

    }
}
