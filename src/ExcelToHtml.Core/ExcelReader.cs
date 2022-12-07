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
    public class ExcelReader
    {
        public async Task<string> Read(string file)
        {
            var ext = Path.GetExtension(file).ToLower();
            if (!ext.Equals(".xlsx") && !ext.Equals("xls"))
                throw new NotSupportedException("Not supported file format！");
            using var fs = new FileStream(file, FileMode.Open, FileAccess.Read);
            using IWorkbook workbook = ext.Equals(".xlsx") ? new XSSFWorkbook(fs) : new HSSFWorkbook(fs);
            var sheets = workbook.NumberOfSheets;
            var table = new StringBuilder();

            for (var sheetIdx = 0; sheetIdx < sheets; sheetIdx++)
            {
                table.Append("<table>");
                var sheet = workbook.GetSheetAt(sheetIdx);

                for (var rowIdx = 0; rowIdx < sheet.LastRowNum; rowIdx++)
                {
                    table.Append("<tr>");
                    var row = sheet.GetRow(rowIdx);
                    if (row is null)
                        continue;

                    foreach (var cell in row.Cells)
                    {
                        table.AppendFormat("<td>{0}</td>", cell?.ToString());
                    }

                    table.Append("</tr>");
                }
                table.Append("</table>");
            }
            return table.ToString();
        }

        public async Task<ConvertResult> Read2(Stream stream, string password = "")
        {
            using var ep = string.IsNullOrWhiteSpace(password) ? new ExcelPackage(stream) : new ExcelPackage(stream, password);

            var result = new ConvertResult();
            var table = new StringBuilder();
            var mergeCell = new HashSet<int>();
            var tableStyle = new CellStyleReader();
            foreach (var sheet in ep.Workbook.Worksheets)
            {
                if (sheet.Dimension is null)
                    continue;
                var sheetDetail = new Sheet
                {
                    Name= sheet.Name
                };
                table.Append("<table class=\"e2h-table\">");
                for (var rowIdx = 1; rowIdx <= sheet.Dimension.Rows; rowIdx++)
                {
                    table.Append("<tr>");
                    //var row = sheet.Row(rowIdx);
                    //if (row is null)
                    //    continue;
                    for (var cellIdx = 1; cellIdx<=sheet.Dimension.Columns; cellIdx++)
                    {

                        var cell = sheet.Cells[rowIdx, cellIdx];
                        if (cell is null)
                        {
                            table.Append("<td></td>");
                            continue;
                        }

                        var value = cell.Text?.Replace("\n", "<br>");
                        if (string.IsNullOrWhiteSpace(value))
                            value = " ";
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

                        if (cell.IsRichText)
                        {

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
                sheetDetail.Html= table.ToString();
                var styles = new StringBuilder();
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

                result.Style = styles.ToString();

                result.Sheets.Add(sheetDetail);
            }

            return result;
        }
    }
}
