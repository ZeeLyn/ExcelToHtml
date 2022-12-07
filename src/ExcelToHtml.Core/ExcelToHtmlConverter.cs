using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtml.Core
{
    public static class ExcelToHtmlConverter
    {
        public static async Task<ConvertResult> ConvertAsync(Stream stream)
        {
            return await ConvertAsync(stream, string.Empty);
        }
        public static async Task<ConvertResult> ConvertAsync(Stream stream, string password)
        {
            return await new ExcelReader().ReaderAsync(stream, password);
        }

        public static async Task<StringBuilder> GenerateUIAsync(Stream stream)
        {
            return await GenerateUIAsync(stream, string.Empty);
        }
        public static async Task<StringBuilder> GenerateUIAsync(Stream stream, string password)
        {
            var result = await ConvertAsync(stream, password);
            var html = new StringBuilder();
            html.AppendFormat("<style>{0}</style>", result.GlobalStyle);
            var tabs = new StringBuilder();
            html.Append("<div class=\"e2h-main\">");
            for (var i = 0; i < result.Sheets.Count; i++)
            {
                html.AppendFormat("<div class=\"e2h-container{1}\" >{0}</div>", result.Sheets[i].Html, i>0 ? " e2h-hide" : "");
                tabs.AppendFormat("<div class=\"e2h-tab-item{1}\">{0}</div>", result.Sheets[i].Name, i==0 ? " e2h-active" : "");
            }
            html.AppendFormat("<div class=\"e2h-tabs\">{0}</div>", tabs);
            html.Append("</div>");
            html.AppendFormat("<script>{0}</script>", result.Script);
            return html;
        }
    }
}
