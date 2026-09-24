// 排序功能驗證：點欄位排序後，整列資料一起搬動，且匯出/存回仍以 ID 正確對應。
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OrcadPartEditor;

class SortTest
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
        string sample = Path.Combine(root, "sample_export.txt");

        var f = new MainForm();
        f.LoadPath(sample);
        var sec = f.FileForTest.Sections.First();
        int valCol = sec.Columns.IndexOf("Value");
        const int idCol = 1;

        // 記錄排序前：每個 ID 對應的 整列 內容（深拷貝）
        var beforeById = sec.Rows.ToDictionary(r => r[idCol], r => (string[])r.Clone());

        // 升冪排序 Value 欄
        f.SortCore(valCol, true);
        sec = f.FileForTest.Sections.First();

        Check(sec.Rows.Count == 7, "排序後：列數不變(7)");

        // 1) 整列一起搬：每列每一格都與排序前同 ID 的那列完全相同
        bool rowsIntact = sec.Rows.All(r =>
        {
            var b = beforeById[r[idCol]];
            return b.Length == r.Length && !b.Where((v, i) => v != r[i]).Any();
        });
        Check(rowsIntact, "排序後：每個 part 整列資料一起移動（ID/Reference/Value… 對齊不變）");

        // 2) 確實有排序：Value 欄為非遞減（用工具的自然排序比較器）
        var cmp = new NaturalComparer();
        var vals = sec.Rows.Select(r => r[valCol] == "<null>" ? "" : r[valCol]).ToList();
        bool nonDecreasing = true;
        for (int i = 1; i < vals.Count; i++)
            if (cmp.Compare(vals[i - 1], vals[i]) > 0) { nonDecreasing = false; break; }
        Check(nonDecreasing, "排序後：Value 欄為升冪（非遞減）");
        Console.WriteLine("    Value 順序: " + string.Join(", ", vals));

        // 3) 降冪切換
        f.SortCore(valCol, false);
        sec = f.FileForTest.Sections.First();
        var vals2 = sec.Rows.Select(r => r[valCol] == "<null>" ? "" : r[valCol]).ToList();
        bool nonIncreasing = true;
        for (int i = 1; i < vals2.Count; i++)
            if (cmp.Compare(vals2[i - 1], vals2[i]) < 0) { nonIncreasing = false; break; }
        Check(nonIncreasing, "降冪排序：Value 欄為非遞增");

        // 4) 存回 .txt 再讀回：結構完好、ID 與整列對應仍正確（匯入 OrCAD 安全）
        string outTxt = Path.Combine(root, "test", "sort_out.txt");
        f.SaveTo(outTxt);
        var reload = MainForm.ParseFile(outTxt);
        var rsec = reload.Sections.First();
        Check(rsec.Rows.Count == 7 && rsec.ColCount == sec.ColCount, "存回再讀：列數/欄數正確");
        // 存檔時 <null> 會轉成空字串，比對時把兩者都正規化為空再比
        Func<string, string> norm = v => v == "<null>" ? "" : (v ?? "");
        bool mapOk = rsec.Rows.All(r =>
        {
            var b = beforeById[r[idCol]];
            return b.Length == r.Length && !b.Where((v, i) => norm(v) != norm(r[i])).Any();
        });
        Check(mapOk, "存回再讀：每個 ID 整列對應仍正確（排序只改列順序＋<null>轉空，OrCAD 以 ID 匯入）");

        // 5) 排序後再編輯不會錯位：透過真正的編輯路徑(CorePaste，等同手動輸入/貼上)
        //    改排序後最上方那列的 Notes，存回後該值必須跟著它的 ID（顯示列↔資料列對齊）
        int notesCol = sec.Columns.IndexOf("Notes");
        if (notesCol >= 0)
        {
            string topId = sec.Rows[0][idCol];
            f.CorePaste("SORTED_EDIT", 0, notesCol);   // 編輯顯示第 0 列（排序後最上方）
            string outTxt2 = Path.Combine(root, "test", "sort_out2.txt");
            f.SaveTo(outTxt2);
            var r2 = MainForm.ParseFile(outTxt2).Sections.First();
            var hit = r2.Rows.First(r => r[idCol] == topId);
            Check(hit[notesCol] == "SORTED_EDIT", "排序後編輯：值正確寫到對應 ID 的列（無錯位）");
        }

        Console.WriteLine();
        Console.WriteLine(fails == 0 ? "==== ALL PASS ====" : ("==== " + fails + " FAILED ===="));
        return fails == 0 ? 0 : 1;
    }
}
