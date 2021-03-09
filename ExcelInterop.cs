using Microsoft.Office.Interop.Excel;
// Solution to the issue "Excel sorting is stuck when ClosedXML is used"
// https://stackoverflow.com/a/66543916/11323942
namespace SortXlsxLib
{
    public static class ExcelInterop
    {
        public static void SortTable(string path, string sheet,
            // start of data, header excluded, e.g row 2 if header is in row 1
            int start_row, int start_col, int end_row, int end_col,
            int sort_col)
        {
            var excelApp = new Application();
            excelApp.Visible = false;

            Workbook wb = excelApp.Workbooks.Open(path);

            Worksheet ws = (Worksheet)wb.Worksheets[sheet];

            var tab = (Range)ws.Range[ws.Cells[start_row, start_col], ws.Cells[end_row, end_col]];

            tab.Sort(ws.Cells[start_row, sort_col], XlSortOrder.xlAscending);

            wb.Save();
            wb.Close();
            excelApp.Quit();
        }
    }
}
