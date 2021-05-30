using System.Collections.Generic;

namespace LegalCitationVSTO
{
    // To communicate between the ribbon and ThisAddIn
    internal static class GlobalState
    {
        public static bool Enabled { get; set; } = true;

        public static Dictionary<int, string> TitleDict { get; } = new Dictionary<int, string>();
    }
}
