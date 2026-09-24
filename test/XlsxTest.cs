// 驗證自帶 xlsx 讀寫的有效性：寫一個 .xlsx → 用 XlsxReader 讀回，比對內容。
using System;
using System.Collections.Generic;
using OrcadPartEditor;

class XlsxTest
{
    static int Main()
    {
        var sec = new PropSection();
        sec.Columns.AddRange(new[] { "HEADER", "ID", "Part Reference", "Value", "Notes & <stuff>" });
        sec.Rows.Add(new[] { "PARTOCC:1", "1", "R1", "10k", "a&b<c>\"d\"" });
        sec.Rows.Add(new[] { "PARTOCC:2", "2", "C1", "0.1uF", "<null>" });
        sec.Rows.Add(new[] { "PARTOCC:3", "3", "U1", "STM32F411", "MCU 微控制器" });

        string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "xlsxtest.xlsx");
        XlsxWriter.Write(path, new List<PropSection> { sec });
        Console.WriteLine("wrote " + path);

        var sheets = XlsxReader.Read(path);
        if (sheets.Count != 1) { Console.WriteLine("FAIL sheet count " + sheets.Count); return 1; }
        var rows = sheets[0];
        Console.WriteLine("rows read = " + rows.Count);
        foreach (var r in rows) Console.WriteLine("  [" + string.Join(" | ", r) + "]");

        // 驗證標題列與資料
        bool ok = true;
        ok &= rows[0].Length >= 5 && rows[0][4] == "Notes & <stuff>";
        ok &= rows[1][2] == "R1" && rows[1][3] == "10k" && rows[1][4] == "a&b<c>\"d\"";
        ok &= rows[3][4] == "MCU 微控制器";  // 空白(<null>)在 xlsx 中為空，這列驗證中文/Unicode
        Console.WriteLine(ok ? "ROUNDTRIP OK" : "ROUNDTRIP FAIL");
        return ok ? 0 : 1;
    }
}
