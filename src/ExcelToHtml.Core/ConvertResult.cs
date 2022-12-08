using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToHtml.Core
{
    public class ConvertResult
    {
        public List<Sheet> Sheets { get; set; } = new();

        public string GlobalStyle { get; set; }

        internal string Script { get; set; }
    }

    public class Sheet
    {
        public string Name { get; set; }
        public string Html { get; set; }
    }
}
