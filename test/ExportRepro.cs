using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OrcadPartEditor;

class ExportRepro
{
    static int Main()
    {
        string root = Directory.GetCurrentDirectory();
        string testDir = Path.Combine(root, "test");
        // 1) GUI 實際路徑：ParseFile -> Sections -> XlsxWriter.Write
        var pf = MainForm.ParseFile(Path.Combine(root, "sample_export.txt"));
        string outp = Path.Combine(testDir, "out_normal.xlsx");
        Try("normal export", () => XlsxWriter.Write(outp, pf.Sections.ToList()));

        // 2) 重複匯出（檔案已存在）
        Try("overwrite export", () => XlsxWriter.Write(outp, pf.Sections.ToList()));

        // 3) 檔案被佔用（模擬 Excel 開著）
        using (var fs = new FileStream(outp, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            Try("locked file export", () => XlsxWriter.Write(outp, pf.Sections.ToList()));
        }

        // 4) 零區塊（檔頭沒有 HEADER 的情況）
        var empty = new List<PropSection>();
        Try("zero-section export", () => XlsxWriter.Write(Path.Combine(testDir, "out_empty.xlsx"), empty));

        // 5) 含特殊字元/中文/空值
        var sec = new PropSection();
        sec.Columns.AddRange(new[] { "HEADER", "ID", "Value", "Note" });
        sec.Rows.Add(new[] { "PARTOCC:1", "1", "0.1uF", "a&b<c>" });
        sec.Rows.Add(new[] { "PARTOCC:2", "2", "<null>", "中文 & \"quote\"" });
        Try("special chars export", () => XlsxWriter.Write(Path.Combine(testDir, "out_special.xlsx"), new List<PropSection> { sec }));

        return 0;
    }

    static void Try(string name, Action a)
    {
        try { a(); Console.WriteLine("OK   : " + name); }
        catch (Exception ex) { Console.WriteLine("FAIL : " + name + "  ->  " + ex.GetType().Name + ": " + ex.Message); }
    }
}
