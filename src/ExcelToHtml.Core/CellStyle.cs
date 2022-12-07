using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToHtml.Core
{
    internal sealed class CellStyle
    {
        internal HashSet<string> ClassNames { get; set; } = new();

        internal string Valign { get; set; }

        internal string Align { get; set; }
    }
}
