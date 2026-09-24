using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OrcadPartEditor;

class UndoTest
{
    static int fails = 0;
    static void Check(bool ok, string name) { Console.WriteLine((ok ? "PASS " : "FAIL ") + name); if (!ok) fails++; }

    [STAThread]
    static int Main()
    {
        Application.EnableVisualStyles();
        var f = new MainForm();
        f.Show();
        Application.DoEvents();
        f.LoadPath(Path.Combine(Directory.GetCurrentDirectory(), "sample_export.txt"));
        Application.DoEvents();
        var sec = f.FileForTest.Sections.First();
        int valCol = sec.Columns.IndexOf("Value");

        Check(f.UndoDepth == 0, "載入後復原堆疊為空");

        string r1v = sec.Rows.First(r => r[2] == "R1")[valCol];   // 原始 R1 Value
        // 操作1：取代
        f.PushUndo(); f.CoreReplace("10k", "ZZZ", true, valCol, null);
        Check(sec.Rows.First(r => r[2] == "R1")[valCol] == "ZZZ", "取代後 R1=ZZZ");
        Check(f.UndoDepth == 1, "堆疊=1");

        // 操作2：條件設定
        f.PushUndo(); f.CoreSetWhere(valCol, -1, false, "AAA", "");
        Check(sec.Rows.All(r => r[valCol] == "AAA"), "全部設成 AAA");
        Check(f.UndoDepth == 2, "堆疊=2");

        // 操作3：新增欄位
        f.PushUndo(); f.CoreAddColumn("XCol", "x");
        Check(sec.Columns.Contains("XCol"), "新增 XCol");
        Check(f.UndoDepth == 3, "堆疊=3");

        // 操作4：再一個（應淘汰最舊，仍=3）
        f.PushUndo(); f.CoreReplace("AAA", "BBB", true, valCol, null);
        Check(f.UndoDepth == 3, "超過上限仍=3 (淘汰最舊)");

        // 復原 4 次
        f.Undo(); Check(sec.Rows.All(r => r[valCol] == "AAA"), "復原1：回到 AAA");
        f.Undo(); Check(!sec.Columns.Contains("XCol"), "復原2：XCol 移除");
        f.Undo(); Check(sec.Rows.Any(r => r[valCol] == "ZZZ"), "復原3：回到 ZZZ 狀態");
        f.Undo(); Check(f.UndoDepth == 0, "第4次：堆疊已空(只存了3步)");
        f.Undo(); // 空堆疊，不應崩潰
        Check(true, "空堆疊復原不崩潰");

        f.Close();
        Console.WriteLine();
        Console.WriteLine(fails == 0 ? "==== UNDO ALL PASS ====" : ("==== " + fails + " FAILED ===="));
        return fails == 0 ? 0 : 1;
    }
}
