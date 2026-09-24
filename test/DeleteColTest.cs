// 刪除欄位驗證：刪掉某屬性欄後，(1) 其餘欄資料不變、(2) 存回 .txt 仍為合法格式、
// (3) 每列欄數與 HEADER 一致、ID 對應不變 → 可正常 Import Properties 匯入 OrCAD。
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OrcadPartEditor;

class DeleteColTest
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
        const int idCol = 1;

        int colsBefore = sec.ColCount;
        int delCol = sec.Columns.IndexOf("Manufacturer");  // 刪一個中間的屬性欄
        Check(delCol >= 2, "找到要刪的欄 Manufacturer");

        // 記錄刪除前：每個 ID 對應的 Value 與 PCB Footprint（驗證其餘欄不受影響）
        int valCol = sec.Columns.IndexOf("Value");
        int fpCol = sec.Columns.IndexOf("PCB Footprint");
        var valById = sec.Rows.ToDictionary(r => r[idCol], r => r[valCol]);
        var fpById = sec.Rows.ToDictionary(r => r[idCol], r => r[fpCol]);

        // 刪欄
        f.CoreDeleteColumn(delCol);
        sec = f.FileForTest.Sections.First();

        Check(sec.ColCount == colsBefore - 1, "刪除後：欄數 -1");
        Check(!sec.Columns.Contains("Manufacturer"), "刪除後：Manufacturer 欄已消失");
        Check(sec.Rows.Count == 7, "刪除後：列數不變(7)");
        Check(sec.Rows.All(r => r.Length == sec.ColCount), "刪除後：每列欄數與標題一致（格式對齊）");

        // 其餘欄資料不變（用刪除後的新欄位索引重新定位）
        int valCol2 = sec.Columns.IndexOf("Value");
        int fpCol2 = sec.Columns.IndexOf("PCB Footprint");
        bool othersIntact = sec.Rows.All(r => r[valCol2] == valById[r[idCol]] && r[fpCol2] == fpById[r[idCol]]);
        Check(othersIntact, "刪除後：其餘欄(Value/PCB Footprint)資料完全不變，且仍對應正確 ID");

        // 存回 .txt 再讀回：格式合法、HEADER 不含被刪欄、每列欄數一致
        string outTxt = Path.Combine(root, "test", "delcol_out.txt");
        f.SaveTo(outTxt);
        var reload = MainForm.ParseFile(outTxt);
        var rsec = reload.Sections.First();
        Check(!rsec.Columns.Contains("Manufacturer"), "存回再讀：HEADER 不再有 Manufacturer 欄");
        Check(rsec.ColCount == colsBefore - 1, "存回再讀：欄數正確(-1)");
        Check(rsec.Rows.Count == 7, "存回再讀：7 列");
        // 解析回來的每列欄數應與 HEADER 相同（OrCAD Import 需要 Tab 欄位對齊）
        Check(rsec.Rows.All(r => r.Length == rsec.ColCount), "存回再讀：每列欄數=HEADER 欄數（匯入 OrCAD 格式正確）");
        // ID 對應 + 其餘欄資料仍正確
        int rval = rsec.Columns.IndexOf("Value");
        bool mapOk = rsec.Rows.All(r => r[rval] == valById[r[idCol]]);
        Check(mapOk, "存回再讀：每個 ID 的 Value 仍正確（刪欄不影響其它屬性匯入）");

        // 不可刪 ID/紀錄鍵欄（防呆）
        int beforeGuard = sec.ColCount;
        f.CoreDeleteColumn(0);
        f.CoreDeleteColumn(1);
        sec = f.FileForTest.Sections.First();
        Check(sec.ColCount == beforeGuard, "防呆：紀錄鍵(0)/ID(1) 欄無法被刪除");

        Console.WriteLine();
        Console.WriteLine(fails == 0 ? "==== ALL PASS ====" : ("==== " + fails + " FAILED ===="));
        return fails == 0 ? 0 : 1;
    }
}
