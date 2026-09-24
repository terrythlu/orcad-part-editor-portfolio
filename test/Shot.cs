using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OrcadPartEditor;

class Shot
{
    [STAThread]
    static int Main()
    {
        Application.EnableVisualStyles();
        var f = new MainForm();
        f.Width = 1200; f.Height = 600;
        f.Show();
        Application.DoEvents();
        string root = Directory.GetCurrentDirectory();
        f.LoadPath(Path.Combine(root, "sample_export.txt"));
        Application.DoEvents();
        System.Threading.Thread.Sleep(400);
        Application.DoEvents();
        var bmp = new Bitmap(f.Width, f.Height);
        f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
        string imageDir = Path.Combine(root, "docs", "images");
        Directory.CreateDirectory(imageDir);
        string outp = Path.Combine(imageDir, "orcad-part-editor.png");
        bmp.Save(outp, System.Drawing.Imaging.ImageFormat.Png);
        Console.WriteLine("saved " + outp);
        f.Close();
        return 0;
    }
}
