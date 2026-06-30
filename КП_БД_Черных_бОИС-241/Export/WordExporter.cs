using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace КП_БД_Черных.Export
{
    public static class WordExporter
    {
        public static void Export(DataGridView dgv)
        {
            if (dgv.Rows.Count == 0) return;
            Word.Application word = new Word.Application();
            word.Documents.Add();
            Word.Document doc = word.ActiveDocument;
            Word.Table table = doc.Tables.Add(doc.Range(), dgv.Rows.Count + 1, dgv.Columns.Count);
            for (int i = 1; i <= dgv.Columns.Count; i++)
                table.Cell(1, i).Range.Text = dgv.Columns[i - 1].HeaderText;
            for (int i = 0; i < dgv.Rows.Count; i++)
                for (int j = 0; j < dgv.Columns.Count; j++)
                    table.Cell(i + 2, j + 1).Range.Text = dgv.Rows[i].Cells[j].Value?.ToString() ?? "";
            word.Visible = true;
        }
    }
}