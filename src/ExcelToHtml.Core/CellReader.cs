using NPOI.SS.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToHtml.Core
{
    public static class CellReader
    {
        public static void ReadCell(this ExcelRange cell)
        {
            var value = cell?.Value?.ToString();
            if (string.IsNullOrWhiteSpace(value))
                value = " ";
            var mergeRows = 0;
            var mergeCols = 0;
            if (cell.Merge)
            {

                //var id = sheet.GetMergeCellId(rowIdx, cellIdx);
                //if (mergeCell.Contains(id))
                //    continue;
                //mergeCell.Add(id);
                //mergeRows = sheet.Cells[sheet.MergedCells[id - 1]].Rows;
                //mergeCols= sheet.Cells[sheet.MergedCells[id - 1]].Columns;
            }

        }
    }
}
