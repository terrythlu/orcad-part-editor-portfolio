// 驗證 OrCAD 屬性檔 parse → serialize 是否結構忠實（包含 <null>、欄位、列順序）。
using System;
using System.IO;
using OrcadPartEditor;

class RoundTrip
{
    static int Main(string[] a)
    {
        string path = a.Length > 0 ? a[0] : Path.Combine(Directory.GetCurrentDirectory(), "sample_export.txt");
        string original = File.ReadAllText(path);

        var pf = MainForm.ParseFile(path);
        string outp = MainForm.Serialize(pf);

        Console.WriteLine("sections=" + System.Linq.Enumerable.Count(pf.Sections));
        foreach (var s in pf.Sections) Console.WriteLine("  cols=" + s.ColCount + " rows=" + s.Rows.Count);

        // 新行為：存檔時 <null> 一律轉成空字串（OrCAD Import 對 <null> 常無法正常更新）。
        // 因此預期輸出 = 原檔把 <null> 換成空字串；其餘結構（欄位、列序、Tab）忠實保留。
        string expected = original.Replace("<null>", "");
        string a1 = expected.Replace("\r\n", "\n").TrimEnd('\n');
        string b1 = outp.Replace("\r\n", "\n").TrimEnd('\n');
        bool identical = a1 == b1;
        bool noNull = !outp.Contains("<null>");
        Console.WriteLine(identical ? "ROUNDTRIP OK (<null>→空字串，結構保留)" : "ROUNDTRIP DIFF");
        Console.WriteLine(noNull ? "NO <null> IN OUTPUT" : "FAIL: 輸出仍含 <null>");
        identical = identical && noNull;
        if (!identical)
        {
            var la = a1.Split('\n'); var lb = b1.Split('\n');
            int n = Math.Max(la.Length, lb.Length);
            for (int i = 0; i < n; i++)
            {
                string x = i < la.Length ? la[i] : "<none>";
                string y = i < lb.Length ? lb[i] : "<none>";
                if (x != y) Console.WriteLine("  line " + (i + 1) + "\n    IN : " + x + "\n    OUT: " + y);
            }
        }

        // 同時驗證「修改一個值後」仍只改該格
        var sec = System.Linq.Enumerable.First(pf.Sections);
        sec.Rows[2][3] = "3.3k";              // R3 value 4.7k -> 3.3k
        string outp2 = MainForm.Serialize(pf);
        bool hasChange = outp2.Contains("3.3k") && !outp2.Contains("\t4.7k\t");
        Console.WriteLine(hasChange ? "EDIT OK" : "EDIT FAIL");

        return (identical && hasChange) ? 0 : 1;
    }
}
