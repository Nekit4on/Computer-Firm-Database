using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace КП_БД_Черных.Export
{
    public static class ExcelExporter
    {
        public static void Export(DataGridView dgv)
        {
            if (dgv.Rows.Count == 0) return;
            Excel.Application excel = new Excel.Application();
            excel.Workbooks.Add();
            Excel.Worksheet sheet = excel.ActiveSheet;
            for (int i = 1; i <= dgv.Columns.Count; i++)
                sheet.Cells[1, i] = dgv.Columns[i - 1].HeaderText;
            for (int i = 0; i < dgv.Rows.Count; i++)
                for (int j = 0; j < dgv.Columns.Count; j++)
                    sheet.Cells[i + 2, j + 1] = dgv.Rows[i].Cells[j].Value?.ToString();
            excel.Visible = true;
        }
    }
}