using System;
using OrcadPartEditor;
class ReadExcel
{
    static int Main(string[] a)
    {
        string path = a.Length > 0 ? a[0] : System.IO.Path.Combine(System.IO.Path.GetTempPath(), "xlsxtest_excel.xlsx");
        var sheets = XlsxReader.Read(path);
        Console.WriteLine("sheets=" + sheets.Count);
        var rows = sheets[0];
        foreach (var r in rows) Console.WriteLine("  [" + string.Join(" | ", r) + "]");
        bool ok = rows[1][2] == "R1" && rows[1][4] == "a&b<c>\"d\"" && rows[3][4] == "MCU 微控制器";
        Console.WriteLine(ok ? "EXCEL-SHAREDSTRINGS READ OK" : "READ FAIL");
        return ok ? 0 : 1;
    }
}
