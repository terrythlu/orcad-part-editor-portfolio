// 端對端測試：建立真正的 MainForm（不顯示視窗），跑過所有使用者功能。
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OrcadPartEditor;

class FullTest
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
        string root = Directory.GetCurrentDirectory();   // 從專案資料夾執行，免受資料夾改名影響
        string dir = Path.Combine(root, "test");
        Directory.CreateDirectory(dir);
        string sample = Path.Combine(root, "sample_export.txt");

        var f = new MainForm();   // 建立但不 Show

        // 1) 載入 + 解析（驗證 #1 表格/標題路徑）
        f.LoadPath(sample);
        var sec = f.FileForTest.Sections.First();
        Check(f.FileForTest.Sections.Count() == 1, "載入：1 個區塊");
        Check(sec.ColCount == 9 && sec.Rows.Count == 7, "載入：9 欄 7 列");
        Check(sec.Columns[0] == "HEADER" && sec.Columns[1] == "ID" && sec.Columns[2] == "Part Reference", "欄位標題正確");

        // 2) 尋找/取代：R0402 -> R0603（限 PCB Footprint 欄=index 4）
        int n1 = f.CoreReplace("R0402", "R0603", true, 4, null);  // R1,R2,R3 三個 R0402
        Check(n1 == 3, "取代：3 個 R0402->R0603");

        // 3) 條件設定：Value 等於 0.1uF 的 PCB Footprint(4) 設成 C0603
        int valCol = sec.Columns.IndexOf("Value");
        int fpCol = sec.Columns.IndexOf("PCB Footprint");
        int n2 = f.CoreSetWhere(fpCol, valCol, false, "C0603", "0.1uF");
        Check(n2 == 2, "條件設定：2 列 0.1uF -> C0603");

        // 4) 新增欄位
        f.CoreAddColumn("Lifecycle", "Active");
        Check(sec.Columns.Contains("Lifecycle") && sec.Rows[0].Last() == "Active", "新增欄位 Lifecycle=Active");

        // 5) Excel 貼上：模擬從 Excel 複製 2x2 區塊，貼到第 0 列、Value 欄
        string clip = "1k\tnewfp\n2k\tnewfp2";
        int n3 = f.CorePaste(clip, 0, valCol);
        Check(n3 == 4, "貼上：4 個儲存格");
        Check(sec.Rows[0][valCol] == "1k" && sec.Rows[1][valCol] == "2k", "貼上：值正確");

        // 5b) 已選取大量儲存格時 Ctrl+V：以選取範圍左上角對齊貼上
        var sel = new HashSet<string>();
        sel.Add("5:" + valCol); sel.Add("5:" + fpCol);
        sel.Add("6:" + valCol); sel.Add("6:" + fpCol);
        int nSel = f.CorePasteIntoSelection("S1\tFPZ1\nS2\tFPZ2", sel);
        Check(nSel == 4, "選取區貼上：更新 2x2");
        Check(sec.Rows[5][valCol] == "S1" && sec.Rows[6][fpCol] == "FPZ2", "選取區貼上：值正確");

        // 5c) 搜尋下一個
        int foundR, foundC;
        bool found = f.CoreFindNext("FPZ2", 0, 0, out foundR, out foundC);
        Check(found && foundR == 6 && foundC == fpCol, "搜尋：找到下一個符合儲存格");

        // 6) 貼上時略過唯讀欄（貼到 ID 欄 index1，應被略過、值不變）
        string idBefore = sec.Rows[0][1];
        string refBefore = sec.Rows[0][2];
        f.CorePaste("ZZZ", 0, 1);
        Check(sec.Rows[0][1] == idBefore, "貼上：ID 欄受保護未被覆蓋");
        Check(sec.Rows[0][2] == refBefore, "貼上：Part Reference 未受影響");

        // 7) 匯出 Excel（驗證 #3）
        string xlsx = Path.Combine(dir, "fulltest.xlsx");
        f.ExportTo(xlsx);
        Check(File.Exists(xlsx) && new FileInfo(xlsx).Length > 0, "匯出 xlsx 成功且非空");

        // 8) 從 Excel 匯入回來（內容應一致，更新 0 格）
        int upd = f.ImportFrom(xlsx);
        Check(upd == 0, "匯入剛匯出的 xlsx：無變化(0格)");
        Check(f.LastImportAnomalies.Count == 0, "匯入掃描：正常 xlsx 無異常符號");

        // 8b) 匯入欄位異常符號掃描：換行、Tab、零寬字元皆須列出
        var anomalySheets = new List<List<string[]>>
        {
            new List<string[]>
            {
                new[] { "HEADER", "ID", "Part Reference", "Value" },
                new[] { "PARTOCC:1", "1", "R1", "line1\r\nline2\tZ\u200B" }
            }
        };
        var anomalies = FieldAnomalyScanner.ScanSheets(anomalySheets);
        Check(anomalies.Any(a => a.Symbol.StartsWith("換行 CRLF") && a.CellAddress == "D2"), "匯入掃描：抓到 CRLF 換行與位置");
        Check(anomalies.Any(a => a.Symbol.StartsWith("Tab")), "匯入掃描：抓到 Tab");
        Check(anomalies.Any(a => a.Symbol.StartsWith("零寬空白")), "匯入掃描：抓到零寬空白");
        Check(anomalies.All(a => a.Preview.IndexOf('\r') < 0 && a.Preview.IndexOf('\n') < 0), "匯入掃描：預覽不含隱藏換行");

        // 8c) 實際寫入含換行的 xlsx，再走完整 XlsxReader + ImportFrom 路徑
        var af = new MainForm();
        af.LoadPath(sample);
        var asec = af.FileForTest.Sections.First();
        int avalCol = asec.Columns.IndexOf("Value");
        var excelSec = new PropSection();
        excelSec.Columns.AddRange(asec.Columns);
        foreach (var row in asec.Rows) excelSec.Rows.Add((string[])row.Clone());
        excelSec.Rows[0][avalCol] = "top\nbottom";
        string anomalyXlsx = Path.Combine(dir, "anomaly_import.xlsx");
        XlsxWriter.Write(anomalyXlsx, new List<PropSection> { excelSec });
        int anomalyUpdated = af.ImportFrom(anomalyXlsx);
        Check(anomalyUpdated == 1 && asec.Rows[0][avalCol] == "top\nbottom", "實際 xlsx 匯入：含換行值已讀入");
        Check(af.LastImportAnomalies.Any(a => a.IsNewLine && a.CellAddress == XlsxWriter.ColLetter(avalCol) + "2"),
            "實際 xlsx 匯入：清單抓到換行與正確儲存格");

        // 9) 改 Excel 內容再匯入：用第二支程式不便，這裡直接驗證 round-trip 對應
        //    存回 .txt 並重新解析，確認結構完好、改動有保留
        string outTxt = Path.Combine(dir, "fulltest_out.txt");
        f.SaveTo(outTxt);
        var reload = MainForm.ParseFile(outTxt);
        var rsec = reload.Sections.First();
        Check(rsec.Columns.Contains("Lifecycle"), "存回再讀：新欄位保留");
        Check(rsec.Rows.First(r => r[2] == "R1")[valCol] == "1k", "存回再讀：R1 Value=1k 保留");
        Check(rsec.Rows.First(r => r[2] == "C1")[fpCol] == "C0603", "存回再讀：C1 footprint=C0603 保留");

        // 10) 後援解析：沒有 HEADER 的純 TSV 檔
        string noHeader = Path.Combine(dir, "noheader.txt");
        File.WriteAllText(noHeader, "Ref\tValue\tFootprint\nR1\t10k\tR0402\nC1\t1uF\tC0402\n");
        var g = new MainForm();
        g.LoadPath(noHeader);
        Check(g.FileForTest.Sections.Any(), "後援：無 HEADER 也能解析成表格");
        var gsec = g.FileForTest.Sections.First();
        Check(gsec.Columns[0] == "Ref" && gsec.Rows.Count == 2, "後援：標題=Ref，2 列");
        // 後援檔也能匯出 xlsx
        string gx = Path.Combine(dir, "noheader.xlsx");
        g.ExportTo(gx);
        Check(File.Exists(gx), "後援：可匯出 xlsx");

        // 11) 多區塊（parts + 第二個 HEADER）
        string multi = Path.Combine(dir, "multi.txt");
        File.WriteAllText(multi,
            "DESIGN\tC:\\my.dsn\n" +
            "HEADER\tID\tPart Reference\tValue\n" +
            "PARTOCC:1\t1\tR1\t10k\n" +
            "HEADER\tID\tNet Name\n" +
            "NETOCC:9\t9\tVCC\n");
        var m = new MainForm();
        m.LoadPath(multi);
        Check(m.FileForTest.Sections.Count() == 2, "多區塊：2 個 HEADER");
        // 存回應位元組保留 DESIGN 行
        string mout = Path.Combine(dir, "multi_out.txt");
        m.SaveTo(mout);
        Check(File.ReadAllText(mout).Contains("DESIGN\tC:\\my.dsn"), "多區塊：DESIGN 行原樣保留");

        Console.WriteLine();
        Console.WriteLine(fails == 0 ? "==== ALL PASS ====" : ("==== " + fails + " FAILED ===="));
        return fails == 0 ? 0 : 1;
    }
}
