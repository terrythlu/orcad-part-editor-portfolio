using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;
using OrcadPartEditor;

class DiagHeaders
{
    [STAThread]
    static int Main()
    {
        Application.EnableVisualStyles();
        var f = new MainForm();
        f.Show();
        Application.DoEvents();
        f.LoadPath(Path.Combine(Directory.GetCurrentDirectory(), "sample_export.txt"));
        Application.DoEvents();

        // 反射取出私有 _grid 來檢查表頭
        var grid = (DataGridView)typeof(MainForm).GetField("_grid", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(f);
        Console.WriteLine("ColumnHeadersVisible = " + grid.ColumnHeadersVisible);
        Console.WriteLine("ColumnHeadersHeight  = " + grid.ColumnHeadersHeight);
        Console.WriteLine("Columns.Count        = " + grid.Columns.Count);
        Console.WriteLine("前 6 個 HeaderText:");
        for (int i = 0; i < Math.Min(6, grid.Columns.Count); i++)
            Console.WriteLine("   [" + i + "] '" + grid.Columns[i].HeaderText + "'  (DataPropertyName=" + grid.Columns[i].DataPropertyName + ")");
        f.Close();
        return 0;
    }
}
