using OfficeOpenXml;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace dotNetUseCase.Helpers
{
    public class ExcelHelper
    {
        public static int GetTotalRowCountByAnyNonNullData(ExcelWorksheet sheet)
        {
            var row = sheet.Dimension.End.Row;
            while (row >= 1)
            {
                var range = sheet.Cells[row, 1, row, sheet.Dimension.End.Column];
                if (range.Any(c => !string.IsNullOrEmpty(c.Text)))
                {
                    break;
                }
                row--;
            }
            return row;
        }

        public static int GetTotalCellCountByAnyNonNullData(ExcelWorksheet sheet)
        {
            var column = sheet.Dimension.End.Column;
            while (column >= 1)
            {
                var range = sheet.Cells[1, column, column, sheet.Dimension.End.Column];
                if (range.Any(c => !string.IsNullOrEmpty(c.Text)))
                {
                    break;
                }
                column--;
            }
            return column;
        }

        public static string MakeValidName(string name)
        {
            string invalidChars = Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()));
            string invalidReStr = string.Format(@"[{0}]+", invalidChars);
            string replace = Regex.Replace(name, invalidReStr, "_").Replace(";", "").Replace(",", "").Replace(" ", "_");
            return replace;
        }

        public static DataTable GetDataTableFromExcel(ExcelWorksheet ws, bool hasHeader = true)
        {
            //https://stackoverflow.com/questions/13396604/excel-to-datatable-using-epplus-excel-locked-for-editing
            DataTable tbl = new DataTable();
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
            {
                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }
            var startRow = hasHeader ? 2 : 1;
            for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                DataRow row = tbl.Rows.Add();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
            }
            return tbl;
        }
    }
}