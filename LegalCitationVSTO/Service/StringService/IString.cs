using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LegalCitationVSTO.Service.StringService
{
    /// <summary>
    /// Interface for StringService.
    /// </summary>
    public interface IString
    {
        /// <summary>
        /// Returns matching substring, or null if absent.
        /// </summary>
        string FindMatch(string text, string pattern);

        /// <summary>
        /// Returns matching footnote text, without sandwiching tokens, or null if absent.
        /// </summary>
        string ExtractFootnoteText(string text);
    }
}
