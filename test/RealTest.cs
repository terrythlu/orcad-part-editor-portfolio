using System;
using System.IO;
using System.Linq;
using System.Text;
using OrcadPartEditor;

class RealTest
{
    static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        if (args.Length == 0 || !File.Exists(args[0]))
        {
            Console.WriteLine("Usage: RealTest.exe <path-to-private-OrCAD-export.txt>");
            return 2;
        }
        string real = Path.GetFullPath(args[0]);

        var pf = MainForm.ParseFile(real);
        Console.WriteLine("偵測編碼: " + pf.Enc.WebName + "  (codepage " + pf.Enc.CodePage + ")");
        Console.WriteLine("Quoted: " + pf.Quoted + "   NewLine: " + (pf.NewLine == "\r\n" ? "CRLF" : "LF"));
        Console.WriteLine("區塊數: " + pf.Sections.Count());
        var s = pf.Sections.First();
        Console.WriteLine("欄數: " + s.ColCount + "   列數: " + s.Rows.Count);
        Console.WriteLine("第0欄='" + s.Columns[0] + "' 第1欄='" + s.Columns[1] + "' 第2欄='" + s.Columns[2] + "'");

        int vCol = s.Columns.IndexOf("Vendor_NO1");
        int dCol = s.Columns.IndexOf("Description");
        var row0 = s.Rows[0];
        Console.WriteLine("Vendor_NO1(欄" + vCol + ") 第1列原值 = '" + row0[vCol] + "'");
        Console.WriteLine("Description(欄" + dCol + ") 第1列 = '" + s.Rows[2][dCol] + "'");

        // 收集所有出現過的非空 Vendor_NO1 值（看中文）
        var vendors = s.Rows.Select(r => r[vCol]).Where(v => v != "<null>" && v.Length > 0)
                            .Distinct().Take(10);
        Console.WriteLine("Vendor_NO1 範例值: " + string.Join(" / ", vendors));

        // round-trip 位元組一致？
        string original = File.ReadAllText(real, pf.Enc);
        string outp = MainForm.Serialize(pf);
        string a = original.Replace("\r\n", "\n").TrimEnd('\n');
        string b = outp.Replace("\r\n", "\n").TrimEnd('\n');
        Console.WriteLine(a == b ? "ROUNDTRIP 位元組一致 ✓" : "ROUNDTRIP 有差異 ✗");
        if (a != b)
        {
            var la = a.Split('\n'); var lb = b.Split('\n');
            for (int i = 0; i < Math.Min(la.Length, lb.Length); i++)
                if (la[i] != lb[i]) { Console.WriteLine("  首個差異在第 " + (i + 1) + " 行");
                    Console.WriteLine("   IN : " + la[i].Substring(0, Math.Min(120, la[i].Length)));
                    Console.WriteLine("   OUT: " + lb[i].Substring(0, Math.Min(120, lb[i].Length))); break; }
            Console.WriteLine("  行數 IN=" + la.Length + " OUT=" + lb.Length);
        }

        // 匯出 xlsx（用真 Excel 驗證中文）
        string xlsx = Path.Combine(Path.GetTempPath(), "OrcadPartEditor-real.xlsx");
        XlsxWriter.Write(xlsx, pf.Sections.ToList());
        Console.WriteLine("已匯出: " + xlsx + "  (" + new FileInfo(xlsx).Length + " bytes)");
        return 0;
    }
}
