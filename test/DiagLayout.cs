using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;
using OrcadPartEditor;

class DiagLayout
{
    [STAThread]
    static int Main()
    {
        Application.EnableVisualStyles();
        var f = new MainForm();
        f.Width = 1100; f.Height = 600;
        f.Show();
        Application.DoEvents();
        f.LoadPath(Path.Combine(Directory.GetCurrentDirectory(), "sample_export.txt"));
        Application.DoEvents();
        System.Threading.Thread.Sleep(200);
        Application.DoEvents();

        var grid = (DataGridView)typeof(MainForm).GetField("_grid", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(f);
        var menu = f.Controls.OfType<MenuStrip>().FirstOrDefault();
        var tool = f.Controls.OfType<ToolStrip>().Where(t => !(t is MenuStrip) && !(t is StatusStrip)).FirstOrDefault();
        var status = f.Controls.OfType<StatusStrip>().FirstOrDefault();

        Console.WriteLine("menu   : Top={0} Bottom={1} H={2}", menu.Top, menu.Bottom, menu.Height);
        Console.WriteLine("tool   : Top={0} Bottom={1} H={2}", tool.Top, tool.Bottom, tool.Height);
        Console.WriteLine("grid   : Top={0} Bottom={1} H={2}", grid.Top, grid.Bottom, grid.Height);
        Console.WriteLine("status : Top={0} Bottom={1}", status.Top, status.Bottom);

        int barsBottom = Math.Max(menu.Bottom, tool.Bottom);
        bool ok = grid.Top >= barsBottom;
        Console.WriteLine("\n選單+工具列底部 = {0}，grid 頂端 = {1}", barsBottom, grid.Top);
        Console.WriteLine(ok ? "LAYOUT OK：grid 在工具列下方，標題列不會被擋" : "LAYOUT BAD：grid 仍被工具列覆蓋 " + (barsBottom - grid.Top) + " px");

        f.Close();
        return ok ? 0 : 1;
    }
}
