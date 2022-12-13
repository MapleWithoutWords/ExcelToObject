using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel
{
    public static class NpoiExtensions
    {
        public static ISheet GetOrCreateSheet(this IWorkbook workbook, string sheetName)
        {
            var sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            return sheet;
        }
        public static IRow GetOrCreateRow(this ISheet sheet, int rowIndex)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                row = sheet.CreateRow(rowIndex);
            }

            return row;
        }
        public static ICell GetOrCreateCell(this IRow row, int cellIndex)
        {
            var cell = row.GetCell(cellIndex);
            if (cell == null)
            {
                cell = row.CreateCell(cellIndex);
            }

            return cell;
        }
    }
}
