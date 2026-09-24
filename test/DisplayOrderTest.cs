// 驗證常用欄位只在畫面前方顯示，儲存與 Undo 仍維持原始欄位順序。
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OrcadPartEditor;

class DisplayOrderTest
{
    static int fails;

    static void Check(bool ok, string name)
    {
        Console.WriteLine((ok ? "PASS " : "FAIL ") + name);
        if (!ok) fails++;
    }

    static string FindColumn(PropSection sec, params string[] names)
    {
        foreach (string name in names)
        {
            string hit = sec.Columns.FirstOrDefault(c => string.Equals(c, name, StringComparison.OrdinalIgnoreCase));
            if (hit != null) return hit;
        }
        return null;
    }

    static void VerifyExternalExp(string path)
    {
        if (!File.Exists(path))
        {
            Check(false, "實際 exp：檔案存在");
            return;
        }

        var f = new MainForm();
        f.LoadPath(path);
        var sec = f.FileForTest.Sections.First();
        string[] originalColumns = sec.Columns.ToArray();
        string item = FindColumn(sec, "Item Number", "Item_Number");
        string manufacturer = FindColumn(sec, "Manufacturer1", "Manufacture1", "MFG1", "MFG");
        string manufacturerPn = FindColumn(sec, "ManufacturerPN1", "ManufacturePN1", "MFGPN1", "MFGPN");
        string[] front = new[] { "Part Reference", "Value", "PCB Footprint", "Description", item, manufacturer, manufacturerPn }
            .Where(c => c != null && sec.Columns.Contains(c)).ToArray();
        var expected = sec.Columns.Take(2).Concat(front)
            .Concat(sec.Columns.Skip(2).Where(c => !front.Contains(c))).ToArray();

        Check(front.Length == 7, "實際 exp：7 個指定欄位均已辨識");
        Check(f.DisplayColumnNamesForTest.SequenceEqual(expected),
            "實際 exp：Part/Value/Footprint/Description/Item/Manufacturer 欄位靠前");
        Check(sec.Columns.SequenceEqual(originalColumns), "實際 exp：原始欄位順序未被改動");

        int valueCol = sec.Columns.IndexOf("Value");
        string edited = (sec.Rows[0][valueCol] ?? "") + "_V23TEST";
        f.CorePaste(edited, 0, valueCol);
        string output = Path.Combine(Path.GetTempPath(), "OrcadPartEditor_v23_display_" + Guid.NewGuid().ToString("N") + ".exp");
        try
        {
            f.SaveTo(output);
            var saved = MainForm.ParseFile(output).Sections.First();
            Check(saved.Columns.SequenceEqual(originalColumns), "實際 exp：儲存仍按原始欄位順序");
            Check(saved.Rows[0][valueCol] == edited, "實際 exp：重排後 Value 編輯未錯位");
        }
        finally
        {
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [STAThread]
    static int Main(string[] args)
    {
        Application.EnableVisualStyles();
        string root = Directory.GetCurrentDirectory();
        string input = Path.Combine(root, "test", "display_order_input.txt");
        string output = Path.Combine(root, "test", "display_order_out.txt");
        File.WriteAllText(input,
            "HEADER\tID\tOtherA\tDescription\tPart Reference\tValue\tOtherB\tPCB Footprint\tItem Number\tManufacturerPN1\tManufacturer1\n" +
            "PARTOCC:1\t1\tA\tD1\tR1\t10k\tB\tR0402\tI1\tPN1\tM1\n" +
            "PARTOCC:2\t2\tA2\tD2\tC1\t0.1uF\tB2\tC0402\tI2\tPN2\tM2\n");

        var f = new MainForm();
        f.LoadPath(input);
        var sec = f.FileForTest.Sections.First();
        string[] originalColumns = sec.Columns.ToArray();
        string[] expectedDisplay =
        {
            "HEADER", "ID", "Part Reference", "Value", "PCB Footprint", "Description",
            "Item Number", "Manufacturer1", "ManufacturerPN1", "OtherA", "OtherB"
        };

        Check(f.DisplayColumnNamesForTest.SequenceEqual(expectedDisplay),
            "畫面：指定欄位依需求集中到最前方（識別欄除外）");
        Check(sec.Columns.SequenceEqual(originalColumns),
            "載入後：底層原始欄位順序未被改動");

        int valueCol = sec.Columns.IndexOf("Value");
        f.CorePaste("20k", 0, valueCol);
        Check(sec.Rows[0][valueCol] == "20k", "畫面重排後編輯：Value 寫回正確原始欄位");

        f.SaveTo(output);
        var saved = MainForm.ParseFile(output).Sections.First();
        Check(saved.Columns.SequenceEqual(originalColumns),
            "儲存：HEADER 按原本欄位順序輸出");
        Check(saved.Rows[0][valueCol] == "20k", "儲存：編輯值保留且未因顯示重排錯位");

        f.PushUndo();
        f.CorePaste("30k", 0, valueCol);
        f.Undo();
        Check(sec.Rows[0][valueCol] == "20k", "Undo：恢復編輯值且不改變原始欄位索引");
        Check(f.DisplayColumnNamesForTest.SequenceEqual(expectedDisplay),
            "Undo：畫面常用欄位順序仍正確");

        string originalXlsx = Path.Combine(root, "test", "display_order_original.xlsx");
        string displayXlsx = Path.Combine(root, "test", "display_order_display.xlsx");
        f.ExportTo(originalXlsx, false);
        f.ExportTo(displayXlsx, true);
        var originalSheet = XlsxReader.Read(originalXlsx).First();
        var displaySheet = XlsxReader.Read(displayXlsx).First();
        Check(originalSheet[0].SequenceEqual(originalColumns),
            "Excel 原始順序選項：欄位順序保持原始 HEADER");
        Check(displaySheet[0].SequenceEqual(expectedDisplay),
            "Excel 畫面順序選項：指定欄位排在前面");

        var imported = new MainForm();
        imported.LoadPath(input);
        imported.ImportFrom(displayXlsx);
        var importedSec = imported.FileForTest.Sections.First();
        Check(importedSec.Rows[0][valueCol] == "20k",
            "Excel 畫面順序回匯：依欄名對應，Value 未錯位");
        imported.RefreshCurrentSectionForTest();
        Check(imported.GridValueForTest(0, valueCol) == "20k",
            "Excel 匯入後畫面：表格內容同步更新");
        if (File.Exists(originalXlsx)) File.Delete(originalXlsx);
        if (File.Exists(displayXlsx)) File.Delete(displayXlsx);

        Check(MainForm.AppVersion == "v2.3", "版本：v2.3");
        if (args.Length > 0) VerifyExternalExp(args[0]);

        Console.WriteLine();
        Console.WriteLine(fails == 0 ? "==== DISPLAY ORDER ALL PASS ====" : ("==== " + fails + " FAILED ===="));
        return fails == 0 ? 0 : 1;
    }
}
