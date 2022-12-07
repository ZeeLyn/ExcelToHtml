using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util.Collections;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtml.Core
{

    internal sealed class ExcelReader
    {
        private string _globalStyle { get; }
        private string _globalScript { get; }

        internal ExcelReader()
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var styleFile = assembly.GetName().Name + ".style.css";
            using var styleStream = assembly.GetManifestResourceStream(styleFile);
            if (styleStream is not null)
            {
                using var reader = new StreamReader(styleStream);
                _globalStyle = reader.ReadToEnd();
            }

            var jsFile = assembly.GetName().Name + ".ui.js";

            using var scriptStream = assembly.GetManifestResourceStream(jsFile);
            if (scriptStream is not null)
            {
                using var reader = new StreamReader(scriptStream);
                _globalScript = reader.ReadToEnd();
            }
        }

        internal async Task<ConvertResult> ReaderAsync(Stream stream, string password = "")
        {
            using var ep = string.IsNullOrWhiteSpace(password) ? new ExcelPackage(stream) : new ExcelPackage(stream, password);
            var result = new ConvertResult
            {
                Script=_globalScript
            };
            var table = new StringBuilder();
            var mergeCell = new HashSet<int>();
            var tableStyle = new CellStyleReader();

            foreach (var sheet in ep.Workbook.Worksheets)
            {
                if (sheet.Dimension is null)
                    continue;

                var sheetDetail = new Sheet
                {
                    Name= sheet.Name,
                };
                table.Append("<table class=\"e2h-table\">");
                for (var rowIdx = 1; rowIdx <= sheet.Dimension.Rows; rowIdx++)
                {
                    var row = sheet.Row(rowIdx);
                    table.AppendFormat("<tr><td{1}>{0}</td>", rowIdx, $" height=\"{(int)row.Height}\"");

                    for (var cellIdx = 1; cellIdx<=sheet.Dimension.Columns; cellIdx++)
                    {
                        var cell = sheet.Cells[rowIdx, cellIdx];
                        if (cell is null)
                        {
                            table.Append("<td></td>");
                            continue;
                        }
                        var value = cell.Text?.Replace("\n", "<br>");
                        //if (value=="7732585"&& rowIdx==46)
                        //    value = value;

                        var mergeRows = 0;
                        var mergeCols = 0;
                        if (cell.Merge)
                        {
                            var id = sheet.GetMergeCellId(rowIdx, cellIdx);
                            //如果添加过td标签，直接跳过
                            if (mergeCell.Contains(id))
                                continue;
                            mergeCell.Add(id);
                            mergeRows = sheet.Cells[sheet.MergedCells[id - 1]].Rows;
                            mergeCols= sheet.Cells[sheet.MergedCells[id - 1]].Columns;
                        }

                        var style = tableStyle.Convert(cell);
                        table.AppendFormat("<td{1}{2}{3}{4}{5}>{0}</td>", value,
                            cell.Merge && mergeRows > 1 ? " rowspan=\"" + mergeRows + "\"" : string.Empty,
                            cell.Merge && mergeCols > 1 ? " colspan=\"" + mergeCols + "\"" : string.Empty,
                            style.ClassNames.Any() ? $" class=\"{string.Join(" ", style.ClassNames)}\"" : string.Empty,
                            string.IsNullOrWhiteSpace(style.Valign) ? string.Empty : $" valign=\"{style.Valign}\"",
                            string.IsNullOrWhiteSpace(style.Align) ? string.Empty : $" align=\"{style.Align}\""
                        );
                    }

                    table.Append("</tr>");
                }
                table.Append("</table>");
                sheetDetail.Html= table;
                var styles = new StringBuilder(_globalStyle);
                foreach (var color in tableStyle.GetAllFontColors())
                {
                    styles.AppendFormat(".e2h-table .{0}{{color:{1}}}", color.Value, color.Key);
                }

                foreach (var color in tableStyle.GetAllFillColors())
                {
                    styles.AppendFormat(".e2h-table .{0}{{background:{1}}}", color.Value, color.Key);
                }

                foreach (var size in tableStyle.GetAllFontSize())
                {
                    styles.AppendFormat(".e2h-table .fs-{0}{{font-size:{0}px}}", size);
                }

                result.GlobalStyle = styles.ToString();

                result.Sheets.Add(sheetDetail);
            }

            return await Task.FromResult(result);
        }
    }
}
