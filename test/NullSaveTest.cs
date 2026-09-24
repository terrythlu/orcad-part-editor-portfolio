// 驗證：存檔時空白 / <null> 一律寫成空字串（引號檔即為 ""），不再輸出 <null>，
// 以避免 OrCAD Import Properties 吃到 <null> 時屬性無法正常清除/更新。
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OrcadPartEditor;

class NullSaveTest
{
    static int fails = 0;
    static void Check(bool ok, string name)
    {
        Console.WriteLine((ok ? "PASS " : "FAIL ") + name);
        if (!ok) fails++;
    }

    [STAThread]
    static int Main()
    {
        Application.EnableVisualStyles();
        string root = Directory.GetCurrentDirectory();

        // 1) 無引號檔（sample_export.txt 內含多個 <null>）
        var f = new MainForm();
        f.LoadPath(Path.Combine(root, "sample_export.txt"));
        string outTxt = Path.Combine(root, "test", "nullsave_out.txt");
        f.SaveTo(outTxt);
        string txt = File.ReadAllText(outTxt);
        Check(!txt.Contains("<null>"), "無引號檔：存檔後整檔不含任何 <null>");

        var rl = MainForm.ParseFile(outTxt).Sections.First();
        int notes = rl.Columns.IndexOf("Notes");
        var r1 = rl.Rows.First(r => r[2] == "R1");     // R1 原本 Notes = <null>
        Check(r1[notes] == "", "無引號檔：原 <null> 欄重讀為空字串");
        var c3 = rl.Rows.First(r => r[2] == "C3");      // C3 Notes = bulk（非空）
        Check(c3[notes] == "bulk", "無引號檔：非空值(bulk)完整保留");

        // 2) 引號檔：空值要寫成 ""（兩個雙引號），不是 <null> 也不是省略欄
        var pf = new PropFile { Quoted = true };
        var sec = new PropSection();
        sec.Columns.AddRange(new[] { "HEADER", "ID", "Part Reference", "Notes" });
        sec.Rows.Add(new[] { "PARTOCC:1", "1", "R1", "<null>" });
        sec.Rows.Add(new[] { "PARTOCC:2", "2", "R2", "hello" });
        pf.Segments.Add(sec);
        string s = MainForm.Serialize(pf);
        Check(!s.Contains("<null>"), "引號檔：序列化結果不含 <null>");
        Check(s.Contains("\"R1\"\t\"\""), "引號檔：空值寫成 \"\"（空引號欄）");
        Check(s.Contains("\"hello\""), "引號檔：非空值保留為 \"hello\"");

        Console.WriteLine();
        Console.WriteLine(fails == 0 ? "==== ALL PASS ====" : ("==== " + fails + " FAILED ===="));
        return fails == 0 ? 0 : 1;
    }
}
