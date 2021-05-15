using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LegalCitationVSTO.Service.StringService
{
    public interface IString
    {
        string FindMatch(string text, string pattern);
        string ExtractFootnoteText(string text);
    }
}
