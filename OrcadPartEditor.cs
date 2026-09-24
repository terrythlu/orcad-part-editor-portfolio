// OrCAD Capture Part Field Editor
// 用途：把 OrCAD Capture「Export Properties」匯出的屬性檔載入成表格，
//       批次快速修改所有 part 的欄位，匯出/匯入 Excel(.xlsx)，再存回給
//       OrCAD「Import Properties」使用（格式完整保留）。
//
// 編譯：用 build.ps1（呼叫 Windows 內建 csc.exe），無須安裝任何東西。
// 相依：純 .NET Framework + WinForms，自帶 xlsx 讀寫，不需要 Excel 或 Python。
//
// 作者注解：OrCAD 匯出檔結構
//   HEADER<tab>ID<tab>Part Reference<tab>Value<tab>...<tab>自訂欄位
//   PARTOCC:6066<tab>6066<tab>D1<tab>3V3<tab><null><tab>...
//   - 第 0 欄是紀錄鍵(HEADER / PARTOCC:n)，第 1 欄是 ID，兩者「不可修改」。
//   - <null> 代表空值；欄位以 Tab 分隔，可能有雙引號。
//   本工具完整保留檔案結構、引號風格、編碼與非資料行，只改你編輯過的儲存格。
// 版本：v2.3

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;

[assembly: AssemblyVersion("2.3.0.0")]
[assembly: AssemblyFileVersion("2.3.0.0")]

namespace OrcadPartEditor
{
    // 一個屬性區塊：一行 HEADER + 其下所有資料列
    class PropSection
    {
        public List<string> Columns = new List<string>();      // 欄名，Columns[0]="HEADER"
        public List<string[]> Rows = new List<string[]>();     // 每列欄位（與 Columns 對齊）
        public string Title;                                    // 顯示用標題

        public int ColCount { get { return Columns.Count; } }
    }

    // 檔案模型：依序保存「原始行(passthrough)」或「屬性區塊」
    class PropFile
    {
        public List<object> Segments = new List<object>();      // string(原始行) 或 PropSection
        public bool Quoted = false;                             // 原檔欄位是否加引號
        public string NewLine = "\r\n";
        public Encoding Enc = Encoding.Default;

        public IEnumerable<PropSection> Sections
        {
            get { foreach (var s in Segments) if (s is PropSection) yield return (PropSection)s; }
        }
    }

    class FieldAnomaly
    {
        public int SectionNumber;
        public string CellAddress;
        public string RecordKey;
        public string ColumnName;
        public string Symbol;
        public int Count;
        public string Preview;
        public bool IsNewLine;
    }

    static class FieldAnomalyScanner
    {
        public static List<FieldAnomaly> ScanSheets(List<List<string[]>> sheets)
        {
            var result = new List<FieldAnomaly>();
            if (sheets == null) return result;

            for (int si = 0; si < sheets.Count; si++)
            {
                var sheet = sheets[si];
                if (sheet == null || sheet.Count == 0) continue;
                string[] header = sheet[0] ?? new string[0];
                for (int r = 0; r < sheet.Count; r++)
                {
                    string[] row = sheet[r] ?? new string[0];
                    string recordKey = r == 0 ? "(欄名)" : (row.Length > 0 ? row[0] ?? "" : "");
                    for (int c = 0; c < row.Length; c++)
                    {
                        string columnName = r == 0
                            ? "(欄位名稱)"
                            : (c < header.Length && !string.IsNullOrEmpty(header[c]) ? header[c] : "第 " + (c + 1) + " 欄");
                        AddValueAnomalies(result, si + 1, XlsxWriter.ColLetter(c) + (r + 1),
                            recordKey, columnName, row[c] ?? "");
                    }
                }
            }
            return result;
        }

        static void AddValueAnomalies(List<FieldAnomaly> result, int sectionNumber, string cellAddress,
            string recordKey, string columnName, string value)
        {
            var found = new Dictionary<string, int>();
            var order = new List<string>();
            for (int i = 0; i < value.Length; i++)
            {
                char ch = value[i];
                string symbol = null;
                if (ch == '\r' && i + 1 < value.Length && value[i + 1] == '\n')
                {
                    symbol = "換行 CRLF (\\r\\n)";
                    i++;
                }
                else if (ch == '\n') symbol = "換行 LF (\\n)";
                else if (ch == '\r') symbol = "換行 CR (\\r)";
                else if (ch == '\t') symbol = "Tab (\\t)";
                else if (ch == '\u2028') symbol = "Unicode 行分隔符 (U+2028)";
                else if (ch == '\u2029') symbol = "Unicode 段落分隔符 (U+2029)";
                else if (ch == '\u00A0') symbol = "不換行空白 NBSP (U+00A0)";
                else if (ch == '\u200B') symbol = "零寬空白 (U+200B)";
                else if (ch == '\u200C') symbol = "零寬非連接符 (U+200C)";
                else if (ch == '\u200D') symbol = "零寬連接符 (U+200D)";
                else if (ch == '\u2060') symbol = "字詞連接符 (U+2060)";
                else if (ch == '\uFEFF') symbol = "BOM / 零寬不換行空白 (U+FEFF)";
                else if (ch == '\uFFFD') symbol = "Unicode 替代字元 (U+FFFD)";
                else if (char.IsHighSurrogate(ch))
                {
                    if (i + 1 < value.Length && char.IsLowSurrogate(value[i + 1])) i++;
                    else symbol = "不完整 Unicode 高代理字元 (U+" + ((int)ch).ToString("X4") + ")";
                }
                else if (char.IsLowSurrogate(ch))
                    symbol = "不完整 Unicode 低代理字元 (U+" + ((int)ch).ToString("X4") + ")";
                else if (char.IsControl(ch))
                    symbol = "控制字元 (U+" + ((int)ch).ToString("X4") + ")";

                if (symbol == null) continue;
                if (!found.ContainsKey(symbol)) { found[symbol] = 0; order.Add(symbol); }
                found[symbol]++;
            }

            foreach (string symbol in order)
            {
                result.Add(new FieldAnomaly
                {
                    SectionNumber = sectionNumber,
                    CellAddress = cellAddress,
                    RecordKey = recordKey,
                    ColumnName = columnName,
                    Symbol = symbol,
                    Count = found[symbol],
                    Preview = VisiblePreview(value),
                    IsNewLine = symbol.StartsWith("換行") || symbol.StartsWith("Unicode 行") || symbol.StartsWith("Unicode 段落")
                });
            }
        }

        static string VisiblePreview(string value)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < value.Length; i++)
            {
                char ch = value[i];
                if (ch == '\r' && i + 1 < value.Length && value[i + 1] == '\n') { sb.Append("[CRLF]"); i++; }
                else if (ch == '\n') sb.Append("[LF]");
                else if (ch == '\r') sb.Append("[CR]");
                else if (ch == '\t') sb.Append("[TAB]");
                else if (char.IsHighSurrogate(ch) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]))
                {
                    sb.Append(ch).Append(value[i + 1]);
                    i++;
                }
                else if (IsInvisibleAnomaly(ch)) sb.Append("[U+").Append(((int)ch).ToString("X4")).Append("]");
                else sb.Append(ch);
                if (sb.Length > 160) { sb.Length = 160; sb.Append("…"); break; }
            }
            return sb.ToString();
        }

        static bool IsInvisibleAnomaly(char ch)
        {
            return char.IsControl(ch) || ch == '\u2028' || ch == '\u2029' || ch == '\u00A0' ||
                ch == '\u200B' || ch == '\u200C' || ch == '\u200D' || ch == '\u2060' ||
                ch == '\uFEFF' || ch == '\uFFFD' || char.IsSurrogate(ch);
        }

        public static string BuildReport(List<FieldAnomaly> anomalies)
        {
            var sb = new StringBuilder();
            sb.AppendLine("區塊\t儲存格\t紀錄鍵\t欄位\t異常符號\t數量\t內容預覽");
            foreach (var a in anomalies)
                sb.Append(a.SectionNumber).Append('\t').Append(a.CellAddress).Append('\t')
                  .Append(a.RecordKey).Append('\t').Append(a.ColumnName).Append('\t')
                  .Append(a.Symbol).Append('\t').Append(a.Count).Append('\t').AppendLine(a.Preview);
            return sb.ToString();
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    class MainForm : Form
    {
        internal const string AppVersion = "v2.3";
        PropFile _file;
        string _path;
        PropSection _active;
        DataTable _table;
        DataGridView _grid;
        ComboBox _sectionBox;
        ToolStripTextBox _searchBox;
        ToolStripStatusLabel _status;
        bool _syncing;
        List<FieldAnomaly> _lastImportAnomalies = new List<FieldAnomaly>();

        // ---- 欄位按鈕 / 排序 / 整欄選取狀態 ----
        Font _hdrFont;                 // 欄位標題字型（畫按鈕用）
        int _sortCol = -1;             // 目前排序欄（-1 = 未排序）
        bool _sortAsc = true;          // 升冪/降冪
        int _selColA = -1, _selColB = -1;  // 目前以欄按鈕選取的欄範圍（highlight 用）
        bool _hdrDrag;                 // 是否正在拖曳欄按鈕做多欄選取
        int _hdrAnchorCol = -1;        // 拖曳起點欄
        int _hdrCurCol = -1;           // 拖曳目前到達欄（避免重複重選）
        const int ArrowZoneW = 22;     // 標題右側「排序箭頭」可點區寬度(px)
        const int DelBtnSize = 14;     // 標題左下角「刪除欄位」小按鈕邊長(px)

        public MainForm()
        {
            Text = "OrCAD Part 欄位編輯器 " + AppVersion;
            Width = 1150; Height = 720;
            StartPosition = FormStartPosition.CenterScreen;
            Icon = SystemIcons.Application;

            _grid = new BufferedGrid
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 38,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText,
                RowHeadersWidth = 56,
                EnableHeadersVisualStyles = false
            };
            // 顯眼的欄位標題列：深藍底、白字、加大加粗（實際繪製改由 CellPainting 處理，做成「按鈕」外觀）
            _hdrFont = new Font(_grid.Font.FontFamily, 9.75f, FontStyle.Bold);
            _grid.ColumnHeadersDefaultCellStyle.Font = _hdrFont;
            _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(31, 78, 148);
            _grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            _grid.CellValueChanged += (s, e) => SyncCellToSection(e.RowIndex, e.ColumnIndex);
            _grid.CellBeginEdit += (s, e) => PushUndo();   // 手動編輯前記錄，供 Ctrl+Z 復原
            _grid.KeyDown += GridKeyDown;
            // 欄位按鈕：自繪標題（欄名 + 右側排序箭頭）、點中間選整欄、拖曳跨欄多選、點箭頭排序
            _grid.CellPainting += GridCellPainting;
            _grid.CellMouseDown += GridHeaderMouseDown;
            _grid.CellMouseMove += GridHeaderMouseMove;
            _grid.CellMouseUp += (s, e) => { _hdrDrag = false; };
            _grid.MouseUp += (s, e) => { _hdrDrag = false; };   // 在格子外放開也要結束拖曳
            // 在列首畫出列號，方便對照
            _grid.RowPostPaint += (s, e) =>
            {
                var g = (DataGridView)s;
                var rect = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, g.RowHeadersWidth - 4, e.RowBounds.Height);
                TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), g.Font, rect,
                    Color.DimGray, TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
            };
            Controls.Add(_grid);

            var statusStrip = new StatusStrip();
            _status = new ToolStripStatusLabel("就緒。請先在 OrCAD Capture：Tools ▸ Export Properties 匯出 .txt，再用『開啟』載入。");
            statusStrip.Items.Add(_status);
            Controls.Add(statusStrip);

            // 重要：先加 Fill(grid) 與 Bottom(status)，再加工具列、最後加選單，
            // docking 的 z-order 才正確，選單/工具列才不會蓋住欄位標題列與第 1 列。
            BuildToolbar();
            BuildMenu();
        }

        // ---------- UI ----------
        void BuildMenu()
        {
            var ms = new MenuStrip();
            var file = new ToolStripMenuItem("檔案(&F)");
            file.DropDownItems.Add("開啟 OrCAD 屬性檔...", null, (s, e) => OpenFile());
            file.DropDownItems.Add("儲存 (回存給 Import Properties)", null, (s, e) => SaveFile(false));
            file.DropDownItems.Add("另存新檔...", null, (s, e) => SaveFile(true));
            file.DropDownItems.Add(new ToolStripSeparator());
            file.DropDownItems.Add("匯出 Excel (.xlsx)...", null, (s, e) => ExportXlsx());
            file.DropDownItems.Add("從 Excel 匯入 (.xlsx)...", null, (s, e) => ImportXlsx());
            file.DropDownItems.Add(new ToolStripSeparator());
            file.DropDownItems.Add("離開", null, (s, e) => Close());
            ms.Items.Add(file);

            var edit = new ToolStripMenuItem("批次編輯(&E)");
            edit.DropDownItems.Add("復原 上一步 (Ctrl+Z)", null, (s, e) => Undo());
            edit.DropDownItems.Add(new ToolStripSeparator());
            edit.DropDownItems.Add("尋找/取代...", null, (s, e) => FindReplaceDialog());
            edit.DropDownItems.Add("向下填滿 (選取區)", null, (s, e) => FillDown());
            edit.DropDownItems.Add("條件設定欄位...", null, (s, e) => SetWhereDialog());
            edit.DropDownItems.Add("新增屬性欄位...", null, (s, e) => AddColumnDialog());
            edit.DropDownItems.Add("刪除目前欄位", null, (s, e) => DeleteCurrentColumn());
            ms.Items.Add(edit);

            var help = new ToolStripMenuItem("說明(&H)");
            help.DropDownItems.Add("使用說明", null, (s, e) => ShowHelp());
            ms.Items.Add(help);

            MainMenuStrip = ms;
            Controls.Add(ms);
        }

        void BuildToolbar()
        {
            var tb = new ToolStrip();
            tb.Items.Add(new ToolStripButton("開啟", null, (s, e) => OpenFile()));
            tb.Items.Add(new ToolStripButton("儲存", null, (s, e) => SaveFile(false)));
            tb.Items.Add(new ToolStripSeparator());
            tb.Items.Add(new ToolStripButton("匯出 Excel", null, (s, e) => ExportXlsx()));
            tb.Items.Add(new ToolStripButton("匯入 Excel", null, (s, e) => ImportXlsx()));
            tb.Items.Add(new ToolStripSeparator());
            tb.Items.Add(new ToolStripButton("↶ 復原", null, (s, e) => Undo()) { ToolTipText = "復原上一步 (Ctrl+Z)，最多 3 步" });
            tb.Items.Add(new ToolStripButton("尋找/取代", null, (s, e) => FindReplaceDialog()));
            tb.Items.Add(new ToolStripButton("向下填滿", null, (s, e) => FillDown()));
            tb.Items.Add(new ToolStripButton("條件設定", null, (s, e) => SetWhereDialog()));
            tb.Items.Add(new ToolStripButton("新增欄位", null, (s, e) => AddColumnDialog()));
            tb.Items.Add(new ToolStripSeparator());
            tb.Items.Add(new ToolStripLabel("搜尋:"));
            _searchBox = new ToolStripTextBox { Width = 180 };
            _searchBox.ToolTipText = "輸入關鍵字後按 Enter 搜尋下一個";
            _searchBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    FindNextFromSearchBox();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            };
            tb.Items.Add(_searchBox);
            tb.Items.Add(new ToolStripButton("下一個", null, (s, e) => FindNextFromSearchBox()) { ToolTipText = "搜尋下一個符合的儲存格" });
            tb.Items.Add(new ToolStripSeparator());
            tb.Items.Add(new ToolStripLabel("區塊:"));
            _sectionBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
            _sectionBox.SelectedIndexChanged += (s, e) =>
            {
                if (_syncing) return;
                if (!_restoring) _undo.Clear();   // 切換區塊清掉復原紀錄（快照是針對單一區塊）
                ShowSection(_sectionBox.SelectedIndex);
            };
            var host = new ToolStripControlHost(_sectionBox);
            tb.Items.Add(host);
            Controls.Add(tb);
        }

        // ---------- 開檔 / 解析 ----------
        void OpenFile()
        {
            using (var d = new OpenFileDialog())
            {
                d.Filter = "OrCAD 屬性檔 (*.txt;*.exp)|*.txt;*.exp|所有檔案 (*.*)|*.*";
                if (d.ShowDialog() != DialogResult.OK) return;
                try
                {
                    LoadPath(d.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("開檔失敗：\n" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        internal PropFile FileForTest { get { return _file; } }

        // 只回報畫面上的欄位順序；_active.Columns 永遠維持原始檔案順序。
        internal List<string> DisplayColumnNamesForTest
        {
            get
            {
                return _grid.Columns.Cast<DataGridViewColumn>()
                    .OrderBy(c => c.DisplayIndex)
                    .Select(c => c.HeaderText)
                    .ToList();
            }
        }

        internal void RefreshCurrentSectionForTest()
        {
            if (_file == null || !_file.Sections.Any()) return;
            ShowSection(Math.Max(0, _sectionBox == null ? 0 : _sectionBox.SelectedIndex));
        }

        internal string GridValueForTest(int row, int column)
        {
            if (_table == null || row < 0 || row >= _table.Rows.Count ||
                column < 0 || column >= _table.Columns.Count) return null;
            return Convert.ToString(_table.Rows[row][column]);
        }

        internal void LoadPath(string path)
        {
            _file = ParseFile(path);
            _path = path;
            _undo.Clear();
            Text = "OrCAD Part 欄位編輯器 " + AppVersion + " - " + Path.GetFileName(_path);
            RefreshSectionBox();
            if (_sectionBox != null && _sectionBox.Items.Count > 0) { _sectionBox.SelectedIndex = 0; ShowSection(0); }
            else if (_file.Sections.Any()) ShowSection(0);
            else SetStatus("檔案不是 Tab 分隔的屬性表格，無法解析。請確認是 OrCAD Export Properties 的輸出。");
        }

        internal static PropFile ParseFile(string path)
        {
            byte[] bytes = File.ReadAllBytes(path);
            Encoding enc = DetectEncoding(bytes);
            string text = enc.GetString(StripBom(bytes, enc));
            string nl = text.Contains("\r\n") ? "\r\n" : (text.Contains("\n") ? "\n" : "\r\n");

            var pf = new PropFile { Enc = enc, NewLine = nl };
            string[] lines = text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            // 去掉因尾端換行造成的最後一個空字串（保留中間空白行）
            int count = lines.Length;
            if (count > 0 && lines[count - 1].Length == 0) count--;

            PropSection cur = null;
            bool quotedDetected = false;
            for (int i = 0; i < count; i++)
            {
                string line = lines[i];
                string[] fields = SplitTabs(line);
                string first = Unquote(fields.Length > 0 ? fields[0] : "");

                if (first == "HEADER")
                {
                    if (line.Contains("\"")) quotedDetected = true;
                    cur = new PropSection();
                    for (int c = 0; c < fields.Length; c++) cur.Columns.Add(Unquote(fields[c]));
                    cur.Title = string.Format("區塊 {0}　({1} 欄)", pf.Sections.Count() + 1, cur.Columns.Count);
                    pf.Segments.Add(cur);
                }
                else if (IsKeywordLine(first) || line.Length == 0)
                {
                    pf.Segments.Add(line); // LIBRARY/DESIGN/TITLE/空行 等 → 原樣保留
                }
                else if (cur != null && fields.Length >= 2 && first.Length > 0)
                {
                    // 視為資料列（補滿/對齊欄位數）
                    var row = new string[cur.ColCount];
                    for (int c = 0; c < cur.ColCount; c++)
                        row[c] = c < fields.Length ? Unquote(fields[c]) : "";
                    cur.Rows.Add(row);
                }
                else
                {
                    pf.Segments.Add(line); // 其他未知行 → 原樣保留
                }
            }
            pf.Quoted = quotedDetected;

            // 後援：若沒有任何 HEADER 區塊，但檔案其實是 Tab 分隔表格，
            // 就把第一筆「多欄」的行當欄位標題，其餘當資料，保證表格一定有內容。
            if (!pf.Sections.Any())
                BuildGenericSection(pf, lines, count, ref quotedDetected);

            return pf;
        }

        // 通用 TSV 後援解析
        static void BuildGenericSection(PropFile pf, string[] lines, int count, ref bool quotedDetected)
        {
            int headerIdx = -1;
            for (int i = 0; i < count; i++)
                if (lines[i].IndexOf('\t') >= 0) { headerIdx = i; break; }
            if (headerIdx < 0) return; // 真的不是表格

            pf.Segments.Clear();
            var sec = new PropSection();
            string[] hf = SplitTabs(lines[headerIdx]);
            if (lines[headerIdx].Contains("\"")) quotedDetected = true;
            foreach (var f in hf) sec.Columns.Add(Unquote(f));
            sec.Title = "資料 (" + sec.ColCount + " 欄)";

            for (int i = 0; i < count; i++)
            {
                if (i == headerIdx) { pf.Segments.Add(sec); continue; }
                if (i < headerIdx) { pf.Segments.Add(lines[i]); continue; }
                string[] ff = SplitTabs(lines[i]);
                if (lines[i].Length == 0 || ff.Length < 2) { pf.Segments.Add(lines[i]); continue; }
                var row = new string[sec.ColCount];
                for (int c = 0; c < sec.ColCount; c++) row[c] = c < ff.Length ? Unquote(ff[c]) : "";
                sec.Rows.Add(row);
            }
            pf.Quoted = quotedDetected;
        }

        static readonly HashSet<string> _keywords = new HashSet<string>(
            new[] { "LIBRARY", "DESIGN", "TITLE", "VIEW", "SCHEMATIC", "PAGE", "FILE" });
        static bool IsKeywordLine(string first)
        {
            // 只有「整行只有 1~2 欄的識別行」才視為關鍵字行，避免誤判資料列
            return _keywords.Contains(first);
        }

        // ---------- 區塊顯示 ----------
        void RefreshSectionBox()
        {
            _syncing = true;
            _sectionBox.Items.Clear();
            int idx = 1;
            foreach (var sec in _file.Sections)
            {
                string refName = sec.Columns.Count > 2 ? sec.Columns[2] : "";
                _sectionBox.Items.Add(string.Format("區塊{0} · {1}列 · {2}欄", idx, sec.Rows.Count, sec.ColCount));
                idx++;
            }
            _syncing = false;
        }

        void ShowSection(int index)
        {
            CommitGrid();
            var list = _file.Sections.ToList();
            if (index < 0 || index >= list.Count) return;
            _active = list[index];

            // 重新繫結資料前先清掉排序箭頭與整欄選取的視覺狀態（排序動作會在繫結後自行重設）
            _sortCol = -1; _selColA = _selColB = -1;

            _table = new DataTable();
            for (int c = 0; c < _active.ColCount; c++)
            {
                string name = _active.Columns[c];
                // DataTable 欄名需唯一
                string unique = name; int dup = 1;
                while (_table.Columns.Contains(unique)) unique = name + "_" + (++dup);
                var col = _table.Columns.Add(unique, typeof(string));
                col.Caption = name;
            }
            foreach (var row in _active.Rows)
            {
                var dr = _table.NewRow();
                for (int c = 0; c < _active.ColCount; c++)
                    dr[c] = NullToBlank(row[c]);
                _table.Rows.Add(dr);
            }

            _syncing = true;
            _grid.DataSource = _table;
            _syncing = false;

            // HeaderText / Columns.Index 仍是原始欄位索引；只調整 DataGridView 的
            // DisplayIndex，因此編輯、Undo、Excel 與 Serialize 都不會改動檔案欄序。
            for (int c = 0; c < _grid.Columns.Count; c++)
            {
                _grid.Columns[c].HeaderText = _active.Columns[c];
                _grid.Columns[c].ToolTipText = _active.Columns[c];
                _grid.Columns[c].Width = c == 0 ? 110 : 130;
                _grid.Columns[c].SortMode = DataGridViewColumnSortMode.NotSortable;
                if (c <= 1) // 第0(紀錄鍵)、第1(ID) 欄唯讀
                {
                    _grid.Columns[c].ReadOnly = true;
                    _grid.Columns[c].DefaultCellStyle.BackColor = Color.FromArgb(238, 238, 238);
                    _grid.Columns[c].DefaultCellStyle.ForeColor = Color.DimGray;
                }
            }
            ApplyPreferredDisplayOrder();
            SetStatus(string.Format("區塊 {0}／{1}：{2} 個 part、{3} 個欄位。灰色欄(紀錄鍵、ID)唯讀。",
                index + 1, list.Count, _active.Rows.Count, _active.ColCount));
        }

        // 這些欄位是最常編輯/檢查的欄位，畫面上排在兩個 OrCAD 識別欄之後。
        // 注意：這裡只改 DataGridView.DisplayIndex，不改 _active.Columns 或每列資料。
        static readonly Dictionary<string, int> PreferredDisplayRanks =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                { "PARTREFERENCE", 0 },
                { "VALUE", 1 },
                { "PCBFOOTPRINT", 2 },
                { "DESCRIPTION", 3 },
                { "ITEMNUMBER", 4 },
                { "MANUFACTURER1", 5 },
                { "MANUFACTURERPN1", 6 },

                // 常見 OrCAD Manufacturer 欄名，視為上面欄位的同義名稱。
                { "MFG", 5 },
                { "MFG1", 5 },
                { "MANUFACTURE1", 5 },
                { "MFGPN", 6 },
                { "MFGPN1", 6 },
                { "MANUFACTUREPN1", 6 }
            };

        static string NormalizeDisplayColumnName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "";
            return new string(name.Where(char.IsLetterOrDigit).ToArray()).ToUpperInvariant();
        }

        static int PreferredDisplayRank(string name)
        {
            int rank;
            return PreferredDisplayRanks.TryGetValue(NormalizeDisplayColumnName(name), out rank) ? rank : -1;
        }

        // 回傳欄位的「畫面/Excel 顯示順序」索引；不會修改原始欄位清單。
        static List<int> BuildPreferredColumnOrder(IList<string> columns)
        {
            if (columns == null) return new List<int>();
            return Enumerable.Range(0, columns.Count)
                .OrderBy(i => i < 2 ? i : (PreferredDisplayRank(columns[i]) >= 0
                    ? 2 + PreferredDisplayRank(columns[i]) : 1000))
                .ThenBy(i => i)
                .ToList();
        }

        void ApplyPreferredDisplayOrder()
        {
            if (_grid == null || _grid.Columns.Count == 0) return;

            // 先解除凍結，避免重排時觸發「Frozen 欄必須在左側」的限制。
            foreach (DataGridViewColumn col in _grid.Columns) col.Frozen = false;

            var ordered = BuildPreferredColumnOrder(_active.Columns)
                .Select(i => _grid.Columns[i])
                .ToList();
            for (int i = 0; i < ordered.Count; i++) ordered[i].DisplayIndex = i;

            // 維持原本 UX：識別欄加上第一個常用欄固定在左側；其餘欄位仍可水平捲動。
            int freezeUpto = Math.Min(3, ordered.Count);
            for (int i = 0; i < freezeUpto; i++) ordered[i].Frozen = true;
        }

        // ---------- 復原 (Ctrl+Z，最多 3 步) ----------
        class Snapshot { public int Section; public List<string> Cols; public List<List<string>> Rows; }
        readonly List<Snapshot> _undo = new List<Snapshot>();
        const int UndoMax = 3;
        bool _restoring;
        internal int UndoDepth { get { return _undo.Count; } }

        internal void PushUndo()
        {
            if (_active == null || _restoring) return;
            CommitGrid(); // 先把目前 grid 內容同步進 _active，再拍快照
            var snap = new Snapshot
            {
                Section = _sectionBox != null ? Math.Max(0, _sectionBox.SelectedIndex) : 0,
                Cols = new List<string>(_active.Columns),
                Rows = _active.Rows.Select(r => new List<string>(r)).ToList()
            };
            _undo.Add(snap);
            while (_undo.Count > UndoMax) _undo.RemoveAt(0);
        }

        internal void Undo()
        {
            if (_undo.Count == 0) { SetStatus("沒有可復原的步驟。"); return; }
            var snap = _undo[_undo.Count - 1];
            _undo.RemoveAt(_undo.Count - 1);

            var list = _file.Sections.ToList();
            if (snap.Section >= list.Count) { SetStatus("無法復原（區塊已變更）。"); return; }
            var sec = list[snap.Section];
            sec.Columns = new List<string>(snap.Cols);
            sec.Rows = snap.Rows.Select(r => r.ToArray()).ToList();

            _restoring = true;
            _table = null;                 // 避免 ShowSection 內的 CommitGrid 覆蓋還原結果
            if (_sectionBox != null && _sectionBox.SelectedIndex != snap.Section && snap.Section < _sectionBox.Items.Count)
                _sectionBox.SelectedIndex = snap.Section;
            ShowSection(snap.Section);
            _restoring = false;
            SetStatus(string.Format("已復原一步。（還可復原 {0} 步）", _undo.Count));
        }

        // 把整個 grid（DataTable）寫回 active section
        void CommitGrid()
        {
            if (_active == null || _table == null) return;
            if (_grid.IsCurrentCellInEditMode) _grid.EndEdit();
            int cols = Math.Min(_active.ColCount, _table.Columns.Count);
            for (int r = 0; r < _table.Rows.Count && r < _active.Rows.Count; r++)
                for (int c = 0; c < cols; c++)
                    _active.Rows[r][c] = BlankToNull(Convert.ToString(_table.Rows[r][c]), c);
        }

        void SyncCellToSection(int r, int c)
        {
            if (_syncing || _active == null || r < 0 || c < 0) return;
            if (r < _active.Rows.Count && c < _active.ColCount)
                _active.Rows[r][c] = BlankToNull(Convert.ToString(_table.Rows[r][c]), c);
        }

        // ---------- 欄位按鈕：自繪標題、點中間選整欄、拖曳跨欄多選、點右側箭頭排序 ----------

        // 自繪每個欄位標題：左側欄名（按鈕本體），右側排序箭頭區
        void GridCellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex != -1 || e.ColumnIndex < 0) return; // 只接管欄位標題列（不含左上角）
            bool selected = IsColumnInSelection(e.ColumnIndex);

            // 按鈕底色（選取中時用較亮的藍）
            using (var bg = new SolidBrush(selected ? Color.FromArgb(21, 101, 192) : Color.FromArgb(31, 78, 148)))
                e.Graphics.FillRectangle(bg, e.CellBounds);
            // 按鈕外框，讓每欄看起來像獨立按鈕
            using (var pen = new Pen(Color.FromArgb(70, 120, 190)))
                e.Graphics.DrawRectangle(pen, e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Width - 1, e.CellBounds.Height - 1);

            // 欄名（左側，太長自動省略）；高度留出底部一條供「刪除欄」小按鈕，避免重疊
            var textRect = new Rectangle(e.CellBounds.Left + 5, e.CellBounds.Top,
                Math.Max(2, e.CellBounds.Width - ArrowZoneW - 7), e.CellBounds.Height - DelBtnSize + 1);
            TextRenderer.DrawText(e.Graphics, _grid.Columns[e.ColumnIndex].HeaderText, _hdrFont, textRect,
                Color.White, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

            // 右側排序箭頭區（與按鈕本體用細線分隔）
            var az = new Rectangle(e.CellBounds.Right - ArrowZoneW, e.CellBounds.Top, ArrowZoneW, e.CellBounds.Height);
            using (var pen = new Pen(Color.FromArgb(70, 120, 190)))
                e.Graphics.DrawLine(pen, az.Left, az.Top + 4, az.Left, az.Bottom - 4);
            DrawSortGlyph(e.Graphics, az, e.ColumnIndex);

            // 左下角「刪除欄位」小按鈕（紅底白叉）；ID/紀錄鍵欄(0,1)受保護，不顯示
            if (e.ColumnIndex >= 2)
            {
                var del = DeleteRect(e.CellBounds);
                using (var b = new SolidBrush(Color.FromArgb(190, 60, 60))) e.Graphics.FillRectangle(b, del);
                using (var p = new Pen(Color.FromArgb(235, 170, 170))) e.Graphics.DrawRectangle(p, del.X, del.Y, del.Width - 1, del.Height - 1);
                using (var pen = new Pen(Color.White, 1.4f))
                {
                    e.Graphics.DrawLine(pen, del.Left + 4, del.Top + 4, del.Right - 5, del.Bottom - 5);
                    e.Graphics.DrawLine(pen, del.Right - 5, del.Top + 4, del.Left + 4, del.Bottom - 5);
                }
            }
            e.Handled = true;
        }

        // 「刪除欄」小按鈕的矩形（標題儲存格內，左下角）
        static Rectangle DeleteRect(Rectangle cell)
        {
            return new Rectangle(cell.Left + 3, cell.Bottom - DelBtnSize - 2, DelBtnSize, DelBtnSize);
        }

        void DrawSortGlyph(Graphics gr, Rectangle az, int col)
        {
            int cx = az.Left + az.Width / 2;
            int cy = az.Top + az.Height / 2;
            if (col == _sortCol)
            {
                // 目前排序欄：白色實心單箭頭（升冪向上、降冪向下）
                Point[] tri = _sortAsc
                    ? new[] { new Point(cx - 5, cy + 3), new Point(cx + 5, cy + 3), new Point(cx, cy - 4) }
                    : new[] { new Point(cx - 5, cy - 3), new Point(cx + 5, cy - 3), new Point(cx, cy + 4) };
                using (var b = new SolidBrush(Color.White)) gr.FillPolygon(b, tri);
            }
            else
            {
                // 未排序欄：暗色上下雙箭頭，暗示「此處可點擊排序」
                using (var b = new SolidBrush(Color.FromArgb(150, 180, 220)))
                {
                    gr.FillPolygon(b, new[] { new Point(cx - 4, cy - 1), new Point(cx + 4, cy - 1), new Point(cx, cy - 6) });
                    gr.FillPolygon(b, new[] { new Point(cx - 4, cy + 1), new Point(cx + 4, cy + 1), new Point(cx, cy + 6) });
                }
            }
        }

        // e.X 是相對於該儲存格左上角的座標；判斷是否落在右側箭頭區
        bool InArrowZone(int colIndex, int x)
        {
            return x >= _grid.Columns[colIndex].Width - ArrowZoneW;
        }

        // e.X/e.Y 相對於標題儲存格；判斷是否落在左下角「刪除欄」小按鈕（0,1 欄不可刪）
        bool InDeleteZone(int colIndex, int x, int y)
        {
            if (colIndex < 2) return false;
            int left = 3, top = _grid.ColumnHeadersHeight - DelBtnSize - 2;
            return x >= left && x <= left + DelBtnSize && y >= top && y <= top + DelBtnSize;
        }

        void GridHeaderMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex != -1 || e.ColumnIndex < 0 || e.Button != MouseButtons.Left) return;
            if (InDeleteZone(e.ColumnIndex, e.X, e.Y))
            {
                DeleteColumnUi(e.ColumnIndex);  // 點左下角紅叉 → 刪除整欄
                return;
            }
            if (InArrowZone(e.ColumnIndex, e.X))
            {
                SortByColumn(e.ColumnIndex);   // 點右側箭頭 → 排序（同欄再點切換升/降冪）
                return;
            }
            // 點按鈕本體 → 選取整欄，並開始拖曳（拖到隔壁欄一起選）
            _hdrDrag = true;
            _hdrAnchorCol = e.ColumnIndex;
            _hdrCurCol = e.ColumnIndex;
            SelectColumnRange(e.ColumnIndex, e.ColumnIndex, true);
        }

        void GridHeaderMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (!_hdrDrag || e.ColumnIndex < 0 || e.ColumnIndex == _hdrCurCol) return;
            _hdrCurCol = e.ColumnIndex;
            SelectColumnRange(_hdrAnchorCol, e.ColumnIndex, false);  // 像 Excel：拖曳擴展整欄選取
        }

        bool IsColumnInSelection(int columnIndex)
        {
            if (_selColA < 0 || _selColB < 0 || columnIndex < 0 || columnIndex >= _grid.ColumnCount) return false;
            int a = _grid.Columns[_selColA].DisplayIndex;
            int b = _grid.Columns[_selColB].DisplayIndex;
            int c = _grid.Columns[columnIndex].DisplayIndex;
            return c >= Math.Min(a, b) && c <= Math.Max(a, b);
        }

        // 依「畫面上的相鄰欄」選取 [a..b]，而不是依原始欄位索引選取。
        // 欄位顯示順序被調整後，拖曳選取仍要符合使用者看到的左右順序。
        void SelectColumnRange(int a, int b, bool setCurrent)
        {
            if (_grid.RowCount == 0 || _grid.ColumnCount == 0) return;
            if (a < 0 || b < 0 || a >= _grid.ColumnCount || b >= _grid.ColumnCount) return;
            var visible = _grid.Columns.Cast<DataGridViewColumn>()
                .OrderBy(c => c.DisplayIndex)
                .ToList();
            int lo = Math.Min(_grid.Columns[a].DisplayIndex, _grid.Columns[b].DisplayIndex);
            int hi = Math.Max(_grid.Columns[a].DisplayIndex, _grid.Columns[b].DisplayIndex);
            if (setCurrent)
            {
                try { _grid.CurrentCell = _grid.Rows[0].Cells[visible[lo].Index]; } catch { }
            }
            _grid.ClearSelection();
            for (int i = lo; i <= hi; i++)
            {
                int c = visible[i].Index;
                for (int r = 0; r < _grid.RowCount; r++)
                    _grid.Rows[r].Cells[c].Selected = true;
            }
            // 狀態仍記錄原始欄位索引，方便排序/刪除/編輯路徑直接使用。
            _selColA = a; _selColB = b;
            // 重畫標題列，讓被選欄的按鈕變亮藍
            _grid.Invalidate(new Rectangle(0, 0, _grid.ClientRectangle.Width, _grid.ColumnHeadersHeight));
        }

        // 點欄位排序箭頭：依該欄排序整張表，並同步搬動底層資料(_active.Rows)以保證匯出/存回正確
        void SortByColumn(int colIndex)
        {
            if (_active == null || _table == null || colIndex < 0 || colIndex >= _active.ColCount) return;
            PushUndo();                                   // 排序前拍快照，可 Ctrl+Z 還原原順序
            bool asc = !(colIndex == _sortCol && _sortAsc); // 同欄再點 → 切換升/降冪
            SortCore(colIndex, asc);
            _sortCol = colIndex; _sortAsc = asc;          // SortCore 內的 ShowSection 會清除，這裡重設
            SelectColumnRange(colIndex, colIndex, true);  // 順便選取被排序欄（也會重畫箭頭）
            SetStatus(string.Format("已依「{0}」{1}排序。匯出/存回會用新順序，OrCAD 仍以 ID 對應，匯入正確。",
                _active.Columns[colIndex], asc ? "升冪▲" : "降冪▼"));
        }

        // 純資料重排（可單元測試）：重排 _active.Rows 後重建顯示表格
        internal void SortCore(int colIndex, bool asc)
        {
            if (_active == null || colIndex < 0 || colIndex >= _active.ColCount) return;
            CommitGrid();                                 // 先把編輯中的內容寫回，避免遺失
            var rows = _active.Rows;
            var keys = new string[rows.Count];
            for (int i = 0; i < rows.Count; i++) keys[i] = CellKey(rows[i], colIndex);
            var order = Enumerable.Range(0, rows.Count).ToList();
            var cmp = new NaturalComparer();
            order.Sort((x, y) =>
            {
                int d = cmp.Compare(keys[x], keys[y]);
                if (!asc) d = -d;
                return d != 0 ? d : x.CompareTo(y);       // 穩定排序：相同值維持原順序
            });
            _active.Rows = order.Select(i => rows[i]).ToList();
            int secIdx = _file.Sections.ToList().IndexOf(_active);
            _table = null;                                // 避免 ShowSection 內 CommitGrid 用舊順序覆蓋
            ShowSection(secIdx < 0 ? 0 : secIdx);
        }

        static string CellKey(string[] row, int c)
        {
            string v = c < row.Length ? row[c] : "";
            return v == "<null>" ? "" : (v ?? "");
        }

        // 從選單刪除「目前所在欄」（以目前儲存格或已選取欄為準）
        void DeleteCurrentColumn()
        {
            if (_active == null) { SetStatus("請先開啟檔案。"); return; }
            int col = _selColA >= 2 ? _selColA : (_grid.CurrentCell != null ? _grid.CurrentCell.ColumnIndex : -1);
            if (col < 2) { SetStatus("請先點選要刪除的欄（紀錄鍵 / ID 欄受保護，不可刪）。"); return; }
            DeleteColumnUi(col);
        }

        // 刪除欄位（UI）：確認後刪除整欄，重建表格。ID/紀錄鍵欄受保護。可 Ctrl+Z 復原。
        void DeleteColumnUi(int colIndex)
        {
            if (_active == null || colIndex < 2 || colIndex >= _active.ColCount) return;
            string name = _active.Columns[colIndex];
            var ans = MessageBox.Show(
                string.Format("確定要刪除欄位「{0}」嗎？\n\n此欄（含所有 part 的該屬性值）會從表格與存檔中移除。\n" +
                              "存回 .txt 後格式仍正確，可照常 Tools ▸ Import Properties 匯入 OrCAD\n" +
                              "（OrCAD 以 ID 對應 part，少了這個屬性欄就單純不更新該屬性）。\n\n" +
                              "紀錄鍵 / ID 欄受保護不會被刪。誤刪可按 Ctrl+Z 復原。",
                    name),
                "刪除欄位", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (ans != DialogResult.OK) return;

            PushUndo();                                   // 刪除前拍快照，可復原
            CoreDeleteColumn(colIndex);
            int secIdx = _file.Sections.ToList().IndexOf(_active);
            _table = null;                                // 避免 ShowSection 內 CommitGrid 用舊欄數覆蓋
            ShowSection(secIdx < 0 ? 0 : secIdx);
            SetStatus(string.Format("已刪除欄位「{0}」。存回 .txt 後仍可正常 Import Properties 匯入 OrCAD。", name));
        }

        // 純資料刪欄（可單元測試）：自欄名與每一列移除第 colIndex 欄，維持欄數對齊
        internal void CoreDeleteColumn(int colIndex)
        {
            if (_active == null || colIndex < 2 || colIndex >= _active.Columns.Count) return;
            CommitGrid();                                 // 先把編輯中的內容寫回
            int oldCount = _active.Columns.Count;
            _active.Columns.RemoveAt(colIndex);
            int newCount = _active.Columns.Count;         // = oldCount - 1
            for (int r = 0; r < _active.Rows.Count; r++)
            {
                var old = _active.Rows[r];
                var nw = new string[newCount];
                int j = 0;
                for (int c = 0; c < oldCount && j < newCount; c++)
                {
                    if (c == colIndex) continue;
                    nw[j++] = c < old.Length ? old[c] : "";
                }
                _active.Rows[r] = nw;
            }
        }

        // ---------- 鍵盤：Excel 貼上 / 清除 ----------
        void GridKeyDown(object sender, KeyEventArgs e)
        {
            if (_table == null) return;
            if (e.Control && e.KeyCode == Keys.Z)
            {
                Undo();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                string text = null;
                try { text = Clipboard.GetText(); } catch { }
                if (!string.IsNullOrEmpty(text) && _grid.CurrentCell != null)
                {
                    PushUndo();
                    int n = _grid.SelectedCells.Count > 1
                        ? CorePasteIntoSelection(text, SelectedCellKeys())
                        : CorePaste(text, _grid.CurrentCell.RowIndex, _grid.CurrentCell.ColumnIndex);
                    _grid.Refresh();
                    SetStatus(string.Format("已從剪貼簿貼上 {0} 個儲存格（唯讀的 ID 欄會自動略過）。", n));
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Delete)
            {
                PushUndo();
                int n = 0;
                _syncing = true;
                foreach (DataGridViewCell cell in _grid.SelectedCells)
                {
                    if (cell.ColumnIndex < 2 || cell.RowIndex < 0) continue;
                    if (!Convert.ToString(_table.Rows[cell.RowIndex][cell.ColumnIndex]).Equals(""))
                    {
                        _table.Rows[cell.RowIndex][cell.ColumnIndex] = "";
                        _active.Rows[cell.RowIndex][cell.ColumnIndex] = BlankToNull("", cell.ColumnIndex);
                        n++;
                    }
                }
                _syncing = false;
                _grid.Refresh();
                SetStatus(string.Format("已清除 {0} 個儲存格。", n));
                e.Handled = true;
            }
        }

        // 從 Excel 複製的內容（Tab 分隔、換行分列）貼到表格，從 (startR,startC) 開始
        internal int CorePaste(string text, int startR, int startC)
        {
            if (_table == null || _active == null || startR < 0 || startC < 0) return 0;
            string[][] cells = ParseClipboardCells(text);
            int n = 0;
            _syncing = true;
            for (int i = 0; i < cells.Length; i++)
            {
                int r = startR + i;
                if (r >= _table.Rows.Count) break;
                string[] cols = cells[i];
                for (int j = 0; j < cols.Length; j++)
                {
                    int c = startC + j;
                    if (c >= _active.ColCount) break;
                    if (c < 2) continue;                 // 不覆蓋紀錄鍵/ID
                    string v = cols[j];
                    _table.Rows[r][c] = v;
                    _active.Rows[r][c] = BlankToNull(v, c);
                    n++;
                }
            }
            _syncing = false;
            return n;
        }

        internal int CorePasteIntoSelection(string text, HashSet<string> selSet)
        {
            if (_table == null || _active == null || selSet == null || selSet.Count == 0) return 0;
            string[][] cells = ParseClipboardCells(text);
            if (cells.Length == 0 || cells[0].Length == 0) return 0;

            int minR = int.MaxValue, minC = int.MaxValue, maxR = -1, maxC = -1;
            foreach (string key in selSet)
            {
                string[] parts = key.Split(':');
                if (parts.Length != 2) continue;
                int r, c;
                if (!int.TryParse(parts[0], out r) || !int.TryParse(parts[1], out c)) continue;
                if (r < 0 || c < 0) continue;
                minR = Math.Min(minR, r);
                minC = Math.Min(minC, c);
                maxR = Math.Max(maxR, r);
                maxC = Math.Max(maxC, c);
            }
            if (maxR < 0 || maxC < 0) return 0;

            bool singleCellClipboard = cells.Length == 1 && cells[0].Length == 1;
            int n = 0;
            _syncing = true;
            for (int r = minR; r <= maxR && r < _table.Rows.Count; r++)
            {
                for (int c = minC; c <= maxC && c < _active.ColCount; c++)
                {
                    if (!selSet.Contains(r + ":" + c) || c < 2) continue;
                    int rr = r - minR;
                    int cc = c - minC;
                    if (!singleCellClipboard && (rr >= cells.Length || cc >= cells[rr].Length)) continue;
                    string v = singleCellClipboard ? cells[0][0] : cells[rr][cc];
                    _table.Rows[r][c] = v;
                    _active.Rows[r][c] = BlankToNull(v, c);
                    n++;
                }
            }
            _syncing = false;
            return n;
        }

        static string[][] ParseClipboardCells(string text)
        {
            if (string.IsNullOrEmpty(text)) return new string[0][];
            string[] lines = text.Replace("\r\n", "\n").Replace("\r", "\n").TrimEnd('\n').Split('\n');
            var cells = new string[lines.Length][];
            for (int i = 0; i < lines.Length; i++) cells[i] = lines[i].Split('\t');
            return cells;
        }

        // ---------- 存檔 ----------
        void SaveFile(bool saveAs)
        {
            if (_file == null) { SetStatus("沒有開啟任何檔案。"); return; }
            CommitGrid();
            string target = _path;
            if (saveAs || string.IsNullOrEmpty(target))
            {
                using (var d = new SaveFileDialog())
                {
                    d.Filter = "OrCAD 屬性檔 (*.txt)|*.txt|所有檔案 (*.*)|*.*";
                    d.FileName = string.IsNullOrEmpty(_path) ? "properties.txt" : Path.GetFileName(_path);
                    if (d.ShowDialog() != DialogResult.OK) return;
                    target = d.FileName;
                }
            }
            try
            {
                WriteFile(target);
                _path = target;
                Text = "OrCAD Part 欄位編輯器 " + AppVersion + " - " + Path.GetFileName(_path);
                SetStatus("已儲存：" + target + "　→　在 OrCAD Capture：Tools ▸ Import Properties 匯入此檔。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("存檔失敗：\n" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void WriteFile(string path)
        {
            File.WriteAllText(path, Serialize(_file), _file.Enc);
        }

        internal void SaveTo(string path) { CommitGrid(); WriteFile(path); }

        internal static string Serialize(PropFile _file)
        {
            var sb = new StringBuilder();
            foreach (var seg in _file.Segments)
            {
                if (seg is string)
                {
                    sb.Append((string)seg);
                    sb.Append(_file.NewLine);
                }
                else
                {
                    var sec = (PropSection)seg;
                    sb.Append(JoinFields(sec.Columns.ToArray(), _file.Quoted, -1));
                    sb.Append(_file.NewLine);
                    foreach (var row in sec.Rows)
                    {
                        // 空值一律寫成空字串（引號檔即為 ""），不再輸出 <null>，
                        // 因為 OrCAD Import 吃 <null> 時常無法正常清除/更新屬性。
                        var outRow = new string[row.Length];
                        for (int c = 0; c < row.Length; c++)
                        {
                            string v = row[c] ?? "";
                            outRow[c] = v == "<null>" ? "" : v;
                        }
                        sb.Append(JoinFields(outRow, _file.Quoted, sec.ColCount));
                        sb.Append(_file.NewLine);
                    }
                }
            }
            return sb.ToString();
        }

        static string JoinFields(string[] fields, bool quoted, int colCount)
        {
            var sb = new StringBuilder();
            for (int c = 0; c < fields.Length; c++)
            {
                if (c > 0) sb.Append('\t');
                string v = fields[c] ?? "";
                if (quoted) sb.Append('"').Append(v.Replace("\"", "\"\"")).Append('"');
                else sb.Append(v);
            }
            return sb.ToString();
        }

        // ---------- 批次編輯 ----------
        void FindReplaceDialog()
        {
            if (_table == null) { SetStatus("請先開啟檔案。"); return; }
            using (var f = new Form())
            {
                f.Text = "尋找 / 取代"; f.Width = 430; f.Height = 260;
                f.FormBorderStyle = FormBorderStyle.FixedDialog; f.StartPosition = FormStartPosition.CenterParent;
                f.MaximizeBox = false; f.MinimizeBox = false;
                var lf = new Label { Text = "尋找：", Left = 15, Top = 18, Width = 60 };
                var tf = new TextBox { Left = 80, Top = 15, Width = 320 };
                var lr = new Label { Text = "取代為：", Left = 15, Top = 52, Width = 60 };
                var tr = new TextBox { Left = 80, Top = 49, Width = 320 };
                var cWhole = new CheckBox { Text = "整格完全相符才取代", Left = 80, Top = 82, Width = 320 };
                var cSel = new CheckBox { Text = "僅作用於選取的儲存格", Left = 80, Top = 108, Width = 320 };
                var cCol = new CheckBox { Text = "僅作用於目前欄", Left = 80, Top = 134, Width = 320 };
                var ok = new Button { Text = "取代", Left = 220, Top = 175, Width = 80, DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "取消", Left = 310, Top = 175, Width = 80, DialogResult = DialogResult.Cancel };
                f.Controls.AddRange(new Control[] { lf, tf, lr, tr, cWhole, cSel, cCol, ok, cancel });
                f.AcceptButton = ok; f.CancelButton = cancel;
                if (f.ShowDialog(this) != DialogResult.OK || tf.Text.Length == 0) return;

                int targetCol = cCol.Checked && _grid.CurrentCell != null ? _grid.CurrentCell.ColumnIndex : -1;
                var selSet = cSel.Checked ? SelectedCellKeys() : null;
                PushUndo();
                int n = CoreReplace(tf.Text, tr.Text, cWhole.Checked, targetCol, selSet);
                _grid.Refresh();
                SetStatus(string.Format("取代完成：共更新 {0} 個儲存格。", n));
            }
        }

        internal int CoreReplace(string find, string repl, bool whole, int targetCol, HashSet<string> selSet)
        {
            if (_table == null || string.IsNullOrEmpty(find)) return 0;
            int n = 0;
            _syncing = true;
            for (int r = 0; r < _table.Rows.Count; r++)
                for (int c = 2; c < _active.ColCount; c++) // 跳過唯讀的 0,1 欄
                {
                    if (targetCol >= 0 && c != targetCol) continue;
                    if (selSet != null && !selSet.Contains(r + ":" + c)) continue;
                    string cur = Convert.ToString(_table.Rows[r][c]);
                    string upd = whole ? (cur == find ? repl : cur) : cur.Replace(find, repl);
                    if (upd != cur) { _table.Rows[r][c] = upd; _active.Rows[r][c] = BlankToNull(upd, c); n++; }
                }
            _syncing = false;
            return n;
        }

        void FillDown()
        {
            if (_table == null || _grid.SelectedCells.Count == 0) { SetStatus("請先選取儲存格。"); return; }
            PushUndo();
            // 依欄分組，找每欄最上方選取列的值，填入該欄其餘選取列
            var byCol = new Dictionary<int, List<int>>();
            foreach (DataGridViewCell cell in _grid.SelectedCells)
            {
                if (cell.ColumnIndex < 2) continue; // 唯讀欄不填
                if (!byCol.ContainsKey(cell.ColumnIndex)) byCol[cell.ColumnIndex] = new List<int>();
                byCol[cell.ColumnIndex].Add(cell.RowIndex);
            }
            int n = 0;
            _syncing = true;
            foreach (var kv in byCol)
            {
                var rows = kv.Value; rows.Sort();
                if (rows.Count < 2) continue;
                string val = Convert.ToString(_table.Rows[rows[0]][kv.Key]);
                for (int i = 1; i < rows.Count; i++)
                {
                    _table.Rows[rows[i]][kv.Key] = val;
                    _active.Rows[rows[i]][kv.Key] = BlankToNull(val, kv.Key);
                    n++;
                }
            }
            _syncing = false;
            _grid.Refresh();
            SetStatus(string.Format("向下填滿：更新 {0} 個儲存格。", n));
        }

        void SetWhereDialog()
        {
            if (_table == null || _active == null) { SetStatus("請先開啟檔案。"); return; }
            using (var f = new Form())
            {
                f.Text = "條件設定欄位"; f.Width = 470; f.Height = 230;
                f.FormBorderStyle = FormBorderStyle.FixedDialog; f.StartPosition = FormStartPosition.CenterParent;
                f.MaximizeBox = false; f.MinimizeBox = false;
                var editable = Enumerable.Range(2, _active.ColCount - 2).Select(i => _active.Columns[i]).ToArray();
                var allCols = Enumerable.Range(0, _active.ColCount).Select(i => _active.Columns[i]).ToArray();

                var l1 = new Label { Text = "設定欄位", Left = 15, Top = 20, Width = 60 };
                var cbSet = new ComboBox { Left = 80, Top = 17, Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
                cbSet.Items.AddRange(editable);
                var l2 = new Label { Text = "= ", Left = 240, Top = 20, Width = 20 };
                var tVal = new TextBox { Left = 260, Top = 17, Width = 170 };

                var l3 = new Label { Text = "條件：欄位", Left = 15, Top = 60, Width = 65 };
                var cbWhere = new ComboBox { Left = 80, Top = 57, Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
                cbWhere.Items.Add("(全部)"); cbWhere.Items.AddRange(allCols); cbWhere.SelectedIndex = 0;
                var cbOp = new ComboBox { Left = 240, Top = 57, Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };
                cbOp.Items.AddRange(new object[] { "等於", "包含" }); cbOp.SelectedIndex = 0;
                var tWhere = new TextBox { Left = 325, Top = 57, Width = 105 };

                var ok = new Button { Text = "套用", Left = 270, Top = 140, Width = 75, DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "取消", Left = 355, Top = 140, Width = 75, DialogResult = DialogResult.Cancel };
                if (editable.Length > 0) cbSet.SelectedIndex = 0;
                f.Controls.AddRange(new Control[] { l1, cbSet, l2, tVal, l3, cbWhere, cbOp, tWhere, ok, cancel });
                f.AcceptButton = ok; f.CancelButton = cancel;
                if (f.ShowDialog(this) != DialogResult.OK || cbSet.SelectedIndex < 0) return;

                int setCol = _active.Columns.IndexOf((string)cbSet.SelectedItem);
                int whereCol = cbWhere.SelectedIndex == 0 ? -1 : _active.Columns.IndexOf((string)cbWhere.SelectedItem);
                bool contains = cbOp.SelectedIndex == 1;
                PushUndo();
                int n = CoreSetWhere(setCol, whereCol, contains, tVal.Text, tWhere.Text);
                _grid.Refresh();
                SetStatus(string.Format("條件設定完成：更新 {0} 列的「{1}」= \"{2}\"。", n, cbSet.SelectedItem, tVal.Text));
            }
        }

        internal int CoreSetWhere(int setCol, int whereCol, bool contains, string newVal, string cond)
        {
            if (_table == null || setCol < 2) return 0;
            int n = 0;
            _syncing = true;
            for (int r = 0; r < _table.Rows.Count; r++)
            {
                bool match = whereCol < 0;
                if (!match)
                {
                    string cv = Convert.ToString(_table.Rows[r][whereCol]);
                    match = contains ? cv.IndexOf(cond, StringComparison.OrdinalIgnoreCase) >= 0 : cv == cond;
                }
                if (match)
                {
                    _table.Rows[r][setCol] = newVal;
                    _active.Rows[r][setCol] = BlankToNull(newVal, setCol);
                    n++;
                }
            }
            _syncing = false;
            return n;
        }

        void AddColumnDialog()
        {
            if (_active == null) { SetStatus("請先開啟檔案。"); return; }
            string name = Prompt("新增屬性欄位", "欄位名稱（例如 Manufacturer Part Number）：", "");
            if (string.IsNullOrEmpty(name)) return;
            if (_active.Columns.Contains(name)) { SetStatus("欄位已存在：" + name); return; }
            string def = Prompt("新增屬性欄位", "預設值（可留空）：", "");
            PushUndo();
            CoreAddColumn(name, def);
            ShowSection(Math.Max(0, _sectionBox.SelectedIndex));
            SetStatus("已新增欄位：" + name + "（OrCAD 會在 Import 時對所有 part 建立此屬性）。");
        }

        internal void CoreAddColumn(string name, string def)
        {
            if (_active == null || _active.Columns.Contains(name)) return;
            _active.Columns.Add(name);
            for (int r = 0; r < _active.Rows.Count; r++)
            {
                var old = _active.Rows[r];
                var nw = new string[old.Length + 1];
                Array.Copy(old, nw, old.Length);
                nw[old.Length] = BlankToNull(def, old.Length);
                _active.Rows[r] = nw;
            }
            // 同步到顯示用 DataTable，避免與 _active 欄數不一致
            if (_table != null)
            {
                string unique = name; int dup = 1;
                while (_table.Columns.Contains(unique)) unique = name + "_" + (++dup);
                _table.Columns.Add(unique, typeof(string));
                _syncing = true;
                for (int r = 0; r < _table.Rows.Count; r++) _table.Rows[r][_table.Columns.Count - 1] = def ?? "";
                _syncing = false;
            }
        }

        // ---------- Excel 匯出 ----------
        void ExportXlsx()
        {
            if (_file == null) { SetStatus("請先開啟檔案。"); return; }
            CommitGrid();
            if (!_file.Sections.Any())
            {
                MessageBox.Show("目前沒有可匯出的資料表格。\n請先用『開啟』載入 OrCAD 匯出的 .txt（表格要有內容）。",
                    "無法匯出", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            using (var d = new SaveFileDialog())
            {
                d.Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx";
                d.FileName = (string.IsNullOrEmpty(_path) ? "parts" : Path.GetFileNameWithoutExtension(_path)) + ".xlsx";
                if (d.ShowDialog() != DialogResult.OK) return;
                bool displayOrder;
                if (!ChooseExcelColumnOrder(out displayOrder)) return;
                try
                {
                    ExportTo(d.FileName, displayOrder);
                    SetStatus("已匯出 Excel：" + d.FileName +
                        (displayOrder ? "（依目前畫面欄位順序）" : "（依原始欄位順序）"));
                }
                catch (IOException)
                {
                    MessageBox.Show("無法寫入：\n" + d.FileName +
                        "\n\n這個檔案可能正在 Excel 中開著，請先關閉它再匯出一次。",
                        "檔案被佔用", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("匯出失敗：\n" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        bool ChooseExcelColumnOrder(out bool displayOrder)
        {
            displayOrder = true;
            using (var f = new Form())
            {
                f.Text = "選擇 Excel 欄位順序";
                f.Width = 470; f.Height = 235;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.StartPosition = FormStartPosition.CenterParent;
                f.MaximizeBox = false; f.MinimizeBox = false;

                var label = new Label
                {
                    Text = "請選擇匯出 Excel 的欄位排列方式：",
                    Left = 18, Top = 18, Width = 420, Height = 25
                };
                var display = new RadioButton
                {
                    Text = "依目前畫面順序（指定欄位會排在前面）",
                    Left = 25, Top = 55, Width = 400, Height = 25,
                    Checked = true
                };
                var original = new RadioButton
                {
                    Text = "依原始檔案順序（保持 OrCAD 匯出欄序）",
                    Left = 25, Top = 88, Width = 400, Height = 25
                };
                var ok = new Button
                {
                    Text = "匯出", Left = 265, Top = 145, Width = 80,
                    DialogResult = DialogResult.OK
                };
                var cancel = new Button
                {
                    Text = "取消", Left = 355, Top = 145, Width = 80,
                    DialogResult = DialogResult.Cancel
                };
                f.Controls.AddRange(new Control[] { label, display, original, ok, cancel });
                f.AcceptButton = ok; f.CancelButton = cancel;
                if (f.ShowDialog(this) != DialogResult.OK) return false;
                displayOrder = display.Checked;
                return true;
            }
        }

        internal void ExportTo(string path)
        {
            ExportTo(path, false);
        }

        // displayOrder=true 時只調整 Excel 欄位順序；原始 _file 與後續 .txt 儲存不變。
        internal void ExportTo(string path, bool displayOrder)
        {
            CommitGrid();
            XlsxWriter.Write(path, BuildExcelSections(displayOrder));
        }

        List<PropSection> BuildExcelSections(bool displayOrder)
        {
            var source = _file.Sections.ToList();
            if (!displayOrder) return source;

            var result = new List<PropSection>();
            foreach (var sec in source)
            {
                var order = BuildPreferredColumnOrder(sec.Columns);
                var copy = new PropSection
                {
                    Title = sec.Title,
                    Columns = order.Select(i => sec.Columns[i]).ToList()
                };
                foreach (var row in sec.Rows)
                    copy.Rows.Add(order.Select(i => i < row.Length ? row[i] : "").ToArray());
                result.Add(copy);
            }
            return result;
        }

        void ImportXlsx()
        {
            if (_file == null) { SetStatus("請先開啟原始 OrCAD 屬性檔，再用 Excel 覆蓋更新內容。"); return; }
            using (var d = new OpenFileDialog())
            {
                d.Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx";
                if (d.ShowDialog() != DialogResult.OK) return;
                try
                {
                    int updated = ImportFrom(d.FileName);
                    ShowSection(Math.Max(0, _sectionBox.SelectedIndex));
                    if (_lastImportAnomalies.Count > 0)
                    {
                        int occurrences = _lastImportAnomalies.Sum(a => a.Count);
                        int newLines = _lastImportAnomalies.Where(a => a.IsNewLine).Sum(a => a.Count);
                        SetStatus(string.Format("已匯入並更新 {0} 格；發現 {1} 個異常符號（換行 {2} 個），請查看清單。",
                            updated, occurrences, newLines));
                        ShowImportAnomalies(_lastImportAnomalies, d.FileName);
                    }
                    else
                        SetStatus(string.Format("已從 Excel 匯入並更新 {0} 個儲存格；未發現異常欄位符號。請按『儲存』再到 OrCAD Import。", updated));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("匯入失敗：\n" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        internal int ImportFrom(string path)
        {
            // 先保存畫面中可能尚未離開編輯模式的內容；匯入後清掉舊 DataTable，
            // 避免 ShowSection 的 CommitGrid 把舊畫面值覆蓋回剛從 Excel 匯入的資料。
            CommitGrid();
            var sheets = XlsxReader.Read(path);
            _lastImportAnomalies = FieldAnomalyScanner.ScanSheets(sheets);
            int updated = MergeSheets(_file.Sections.ToList(), sheets);
            _table = null;
            return updated;
        }

        internal List<FieldAnomaly> LastImportAnomalies { get { return _lastImportAnomalies; } }

        void ShowImportAnomalies(List<FieldAnomaly> anomalies, string sourcePath)
        {
            using (var f = new Form())
            {
                f.Text = "匯入異常欄位符號";
                f.Width = 1120; f.Height = 600;
                f.StartPosition = FormStartPosition.CenterParent;
                f.MinimizeBox = false;

                int occurrences = anomalies.Sum(a => a.Count);
                int newLines = anomalies.Where(a => a.IsNewLine).Sum(a => a.Count);
                int cells = anomalies.Select(a => a.SectionNumber + ":" + a.CellAddress).Distinct().Count();
                var info = new Label
                {
                    Dock = DockStyle.Top,
                    Height = 58,
                    Padding = new Padding(10, 8, 10, 4),
                    Text = string.Format("檔案：{0}\n找到 {1} 個異常符號，分布於 {2} 個儲存格；其中換行 {3} 個。資料已匯入，請先檢查再儲存給 OrCAD。",
                        Path.GetFileName(sourcePath), occurrences, cells, newLines)
                };

                var grid = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText,
                    RowHeadersVisible = false
                };
                grid.Columns.Add("Section", "區塊"); grid.Columns[0].Width = 55;
                grid.Columns.Add("Cell", "Excel 儲存格"); grid.Columns[1].Width = 90;
                grid.Columns.Add("Key", "紀錄鍵"); grid.Columns[2].Width = 145;
                grid.Columns.Add("Column", "欄位"); grid.Columns[3].Width = 180;
                grid.Columns.Add("Symbol", "異常符號"); grid.Columns[4].Width = 225;
                grid.Columns.Add("Count", "數量"); grid.Columns[5].Width = 55;
                grid.Columns.Add("Preview", "內容預覽（異常符號已轉成標記）"); grid.Columns[6].Width = 330;
                foreach (var a in anomalies)
                {
                    int index = grid.Rows.Add(a.SectionNumber, a.CellAddress, a.RecordKey,
                        a.ColumnName, a.Symbol, a.Count, a.Preview);
                    if (a.IsNewLine)
                    {
                        grid.Rows[index].DefaultCellStyle.BackColor = Color.MistyRose;
                        grid.Rows[index].DefaultCellStyle.ForeColor = Color.DarkRed;
                    }
                }

                var bottom = new Panel { Dock = DockStyle.Bottom, Height = 48 };
                var copy = new Button { Text = "複製全部清單", Width = 120, Height = 28, Left = 10, Top = 10 };
                copy.Click += (s, e) =>
                {
                    try
                    {
                        Clipboard.SetText(FieldAnomalyScanner.BuildReport(anomalies));
                        copy.Text = "已複製";
                    }
                    catch (Exception ex) { MessageBox.Show("複製失敗：\n" + ex.Message, "錯誤"); }
                };
                var close = new Button { Text = "關閉", Width = 90, Height = 28, Top = 10, Anchor = AnchorStyles.Top | AnchorStyles.Right };
                close.Left = f.ClientSize.Width - close.Width - 28;
                close.DialogResult = DialogResult.OK;
                bottom.Controls.Add(copy); bottom.Controls.Add(close);
                f.AcceptButton = close;
                f.Controls.Add(grid); f.Controls.Add(info); f.Controls.Add(bottom);
                f.ShowDialog(this);
            }
        }

        // 用 Excel 內容更新 section：以「紀錄鍵(第0欄)」對應列，依欄名更新可編輯欄；
        // Excel 若有新欄位則對該 section 新增。ID/紀錄鍵不變。
        int MergeSheets(List<PropSection> sections, List<List<string[]>> sheets)
        {
            int updated = 0;
            for (int si = 0; si < sections.Count && si < sheets.Count; si++)
            {
                var sec = sections[si];
                var sheet = sheets[si];
                if (sheet.Count < 1) continue;
                string[] head = sheet[0];

                // Excel 欄名 → section 欄索引（沒有的就新增欄位）
                var colMap = new int[head.Length];
                for (int hc = 0; hc < head.Length; hc++)
                {
                    string hn = head[hc] ?? "";
                    int idx = sec.Columns.IndexOf(hn);
                    if (idx < 0 && hn.Length > 0 && hc >= 2)
                    {
                        idx = sec.Columns.Count;
                        sec.Columns.Add(hn);
                        for (int r = 0; r < sec.Rows.Count; r++)
                        {
                            var old = sec.Rows[r]; var nw = new string[old.Length + 1];
                            Array.Copy(old, nw, old.Length); nw[old.Length] = "";
                            sec.Rows[r] = nw;
                        }
                    }
                    colMap[hc] = idx;
                }
                // 以第0欄(紀錄鍵)建立索引
                var keyToRow = new Dictionary<string, int>();
                for (int r = 0; r < sec.Rows.Count; r++)
                {
                    string k = sec.Rows[r][0] ?? "";
                    if (!keyToRow.ContainsKey(k)) keyToRow[k] = r;
                }
                for (int sr = 1; sr < sheet.Count; sr++)
                {
                    string[] vals = sheet[sr];
                    if (vals.Length == 0) continue;
                    string key = vals[0] ?? "";
                    int rowIdx;
                    if (!keyToRow.TryGetValue(key, out rowIdx)) continue; // 找不到對應 part 就略過
                    for (int hc = 2; hc < vals.Length && hc < colMap.Length; hc++) // 不動 0,1 欄
                    {
                        int dst = colMap[hc];
                        if (dst < 2) continue;
                        string nv = BlankToNull(vals[hc] ?? "", dst);
                        if (sec.Rows[rowIdx][dst] != nv) { sec.Rows[rowIdx][dst] = nv; updated++; }
                    }
                }
            }
            return updated;
        }

        // ---------- 小工具 ----------
        HashSet<string> SelectedCellKeys()
        {
            var set = new HashSet<string>();
            foreach (DataGridViewCell c in _grid.SelectedCells) set.Add(c.RowIndex + ":" + c.ColumnIndex);
            return set;
        }

        void FindNextFromSearchBox()
        {
            if (_table == null) { SetStatus("請先開啟檔案。"); return; }
            string term = _searchBox == null ? "" : _searchBox.Text;
            if (string.IsNullOrEmpty(term)) { SetStatus("請輸入搜尋文字。"); return; }

            int startR = 0, startC = -1;
            if (_grid.CurrentCell != null)
            {
                startR = _grid.CurrentCell.RowIndex;
                startC = _grid.CurrentCell.ColumnIndex;
            }

            int foundR, foundC;
            if (!CoreFindNext(term, startR, startC, out foundR, out foundC))
            {
                SetStatus("找不到：" + term);
                return;
            }

            _grid.ClearSelection();
            _grid.CurrentCell = _grid.Rows[foundR].Cells[foundC];
            _grid.Rows[foundR].Cells[foundC].Selected = true;
            if (foundR >= 0) _grid.FirstDisplayedScrollingRowIndex = foundR;
            if (foundC >= 0 && !_grid.Columns[foundC].Frozen) _grid.FirstDisplayedScrollingColumnIndex = foundC;
            SetStatus(string.Format("找到「{0}」：第 {1} 列，欄位「{2}」。", term, foundR + 1, _active.Columns[foundC]));
        }

        internal bool CoreFindNext(string term, int startR, int startC, out int foundR, out int foundC)
        {
            foundR = -1;
            foundC = -1;
            if (_table == null || _active == null || string.IsNullOrEmpty(term)) return false;
            int rows = _table.Rows.Count;
            int cols = _active.ColCount;
            if (rows == 0 || cols == 0) return false;

            int total = rows * cols;
            int start = Math.Max(0, Math.Min(total - 1, startR * cols + startC));
            for (int step = 1; step <= total; step++)
            {
                int pos = (start + step) % total;
                int r = pos / cols;
                int c = pos % cols;
                string cur = Convert.ToString(_table.Rows[r][c]);
                if (cur.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    foundR = r;
                    foundC = c;
                    return true;
                }
            }
            return false;
        }

        void ShowHelp()
        {
            MessageBox.Show(
                "操作流程：\n\n" +
                "1) OrCAD Capture 中選取設計，Tools ▸ Export Properties，存成 .txt。\n" +
                "2) 本工具『開啟』該 .txt。灰色的『紀錄鍵 / ID』欄唯讀，其餘可編輯。\n" +
                "3) 欄位按鈕（像 Excel 操作欄）：\n" +
                "     · 點欄位標題中間 → 選取整欄；按住拖到隔壁欄 → 連續多欄一起選\n" +
                "     · 點標題右側的上下箭頭 → 依該欄排序；同一欄再點一次 → 升冪/降冪切換\n" +
                "     · 點標題左下角的紅色叉 → 刪除整欄（ID/紀錄鍵欄受保護；可 Ctrl+Z 復原）\n" +
                "     · 排序/刪欄後仍可正常匯出/存回，OrCAD 以 ID 對應 part，匯入結果正確\n" +
                "     · Part Reference、Value、PCB Footprint、Description、Item Number、Manufacturer1、ManufacturerPN1\n" +
                "       會優先顯示在前面；這只改畫面順序，儲存/Excel/復原仍依原始欄位順序處理\n" +
                "4) 直接在表格修改，或用批次工具：\n" +
                "     · 從 Excel 複製儲存格 → 點本表格起始格 → Ctrl+V 直接貼上\n" +
                "     · Delete 鍵：清空選取的儲存格\n" +
                "     · Ctrl+Z：復原上一步（最多 3 步）\n" +
                "     · 尋找/取代（可限定欄或選取區）\n" +
                "     · 向下填滿（先選取一欄多列）\n" +
                "     · 條件設定欄位（例如：把 Value=0.1uF 的 PCB Footprint 全設成 C0402）\n" +
                "     · 新增屬性欄位（替所有 part 加一個新屬性）\n" +
                "   （左側有列號；前三欄已凍結，左右捲動時仍看得到 Reference）\n" +
                "5) 可『匯出 Excel』給別人看／在 Excel 改完再『匯入 Excel』。\n" +
                "6) 按『儲存』寫回 .txt → 回 OrCAD：Tools ▸ Import Properties 匯入。\n\n" +
                "注意：\n" +
                "  · 請勿更動 ID 欄，OrCAD 靠它對應回每個 part。\n" +
                "  · 空白儲存格存檔時一律寫成空字串(\"\")，不再寫 <null>，\n" +
                "    以免 OrCAD Import 吃到 <null> 時屬性無法正常清除/更新。",
                "使用說明", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void SetStatus(string s) { _status.Text = s; }

        static string Prompt(string title, string label, string def)
        {
            using (var f = new Form())
            {
                f.Text = title; f.Width = 460; f.Height = 160;
                f.FormBorderStyle = FormBorderStyle.FixedDialog; f.StartPosition = FormStartPosition.CenterParent;
                f.MaximizeBox = false; f.MinimizeBox = false;
                var l = new Label { Text = label, Left = 15, Top = 15, Width = 420 };
                var t = new TextBox { Left = 15, Top = 45, Width = 415, Text = def };
                var ok = new Button { Text = "確定", Left = 270, Top = 80, Width = 75, DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "取消", Left = 355, Top = 80, Width = 75, DialogResult = DialogResult.Cancel };
                f.Controls.AddRange(new Control[] { l, t, ok, cancel });
                f.AcceptButton = ok; f.CancelButton = cancel;
                return f.ShowDialog() == DialogResult.OK ? t.Text : null;
            }
        }

        // <null> ↔ 空白：顯示時 <null> 視為空；內部仍以 <null> 當「空值」記號（沿用既有批次邏輯）。
        // 注意：實際「存檔」時會在 Serialize 把 <null> 一律改寫成空字串("")，
        //       因為 OrCAD Import Properties 吃到 <null> 常無法正常清除/更新屬性。
        static string NullToBlank(string v) { return v == "<null>" ? "" : v; }
        static string BlankToNull(string v, int colIndex)
        {
            if (colIndex <= 1) return v ?? "";          // 紀錄鍵/ID 保持原樣
            return string.IsNullOrEmpty(v) ? "<null>" : v;
        }

        static string[] SplitTabs(string line) { return line.Split('\t'); }
        static string Unquote(string s)
        {
            if (s.Length >= 2 && s[0] == '"' && s[s.Length - 1] == '"')
                return s.Substring(1, s.Length - 2).Replace("\"\"", "\"");
            return s;
        }

        static Encoding DetectEncoding(byte[] b)
        {
            if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) return new UTF8Encoding(true);
            if (b.Length >= 2 && b[0] == 0xFF && b[1] == 0xFE) return Encoding.Unicode;
            if (b.Length >= 2 && b[0] == 0xFE && b[1] == 0xFF) return Encoding.BigEndianUnicode;
            // 嘗試嚴格 UTF-8；失敗則用系統 ANSI（OrCAD 常用）
            try { new UTF8Encoding(false, true).GetString(b); return new UTF8Encoding(false); }
            catch { return Encoding.Default; }
        }
        static byte[] StripBom(byte[] b, Encoding enc)
        {
            byte[] pre = enc.GetPreamble();
            if (pre.Length > 0 && b.Length >= pre.Length)
            {
                bool m = true;
                for (int i = 0; i < pre.Length; i++) if (b[i] != pre[i]) { m = false; break; }
                if (m) { var r = new byte[b.Length - pre.Length]; Array.Copy(b, pre.Length, r, 0, r.Length); return r; }
            }
            return b;
        }
    }

    // 開啟雙緩衝的 DataGridView，拖曳選取/重畫標題時不閃爍
    class BufferedGrid : DataGridView
    {
        public BufferedGrid() { DoubleBuffered = true; }
    }

    // 自然排序比較：兩者皆為數字時比數值（避免 "10" 排在 "2" 前），否則比字串；空值排最後
    class NaturalComparer : System.Collections.Generic.IComparer<string>
    {
        public int Compare(string a, string b)
        {
            a = a ?? ""; b = b ?? "";
            double da, db;
            bool na = TryNum(a, out da), nb = TryNum(b, out db);
            if (na && nb)
            {
                int d = da.CompareTo(db);
                return d != 0 ? d : string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
            }
            if (a.Length == 0 && b.Length > 0) return 1;   // 空字串排最後
            if (b.Length == 0 && a.Length > 0) return -1;
            if (na != nb) return na ? -1 : 1;              // 數字排在文字之前
            return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
        }
        static bool TryNum(string s, out double v)
        {
            return double.TryParse(s, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out v);
        }
    }

    // ============ 極簡 XLSX 寫入（inlineStr，全部當文字，避免 Excel 自動轉型）============
    static class XlsxWriter
    {
        public static void Write(string path, List<PropSection> sections)
        {
            if (sections == null || sections.Count == 0)
                throw new InvalidOperationException("沒有可匯出的資料表格（0 個區塊）。");

            // 先寫到暫存檔，成功後再置換目標檔，避免半成品壞檔；
            // 若目標檔被 Excel 開著，置換時丟 IOException，由呼叫端友善提示。
            string tmp = path + ".tmp" + Guid.NewGuid().ToString("N").Substring(0, 6);
            using (var zip = ZipFile.Open(tmp, ZipArchiveMode.Create))
            {
                AddEntry(zip, "[Content_Types].xml", ContentTypes(sections.Count));
                AddEntry(zip, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>");
                AddEntry(zip, "xl/workbook.xml", Workbook(sections.Count));
                AddEntry(zip, "xl/_rels/workbook.xml.rels", WorkbookRels(sections.Count));
                for (int i = 0; i < sections.Count; i++)
                    AddEntry(zip, "xl/worksheets/sheet" + (i + 1) + ".xml", Sheet(sections[i]));
            }
            try
            {
                if (File.Exists(path)) File.Delete(path);
                File.Move(tmp, path);
            }
            catch
            {
                try { File.Delete(tmp); } catch { }
                throw;
            }
        }

        static void AddEntry(ZipArchive zip, string name, string content)
        {
            var e = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (var w = new StreamWriter(e.Open(), new UTF8Encoding(false))) w.Write(content);
        }

        static string ContentTypes(int n)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            for (int i = 0; i < n; i++)
                sb.Append("<Override PartName=\"/xl/worksheets/sheet" + (i + 1) + ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            sb.Append("</Types>");
            return sb.ToString();
        }

        static string Workbook(int n)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
            sb.Append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
            for (int i = 0; i < n; i++)
                sb.Append("<sheet name=\"Section" + (i + 1) + "\" sheetId=\"" + (i + 1) + "\" r:id=\"rId" + (i + 1) + "\"/>");
            sb.Append("</sheets></workbook>");
            return sb.ToString();
        }

        static string WorkbookRels(int n)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 0; i < n; i++)
                sb.Append("<Relationship Id=\"rId" + (i + 1) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet" + (i + 1) + ".xml\"/>");
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        static string Sheet(PropSection sec)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
            // 標題列
            sb.Append(RowXml(1, sec.Columns.ToArray(), true));
            // 資料列
            for (int r = 0; r < sec.Rows.Count; r++)
            {
                var disp = new string[sec.ColCount];
                for (int c = 0; c < sec.ColCount; c++)
                {
                    string v = c < sec.Rows[r].Length ? sec.Rows[r][c] : "";
                    disp[c] = v == "<null>" ? "" : v;
                }
                sb.Append(RowXml(r + 2, disp, false));
            }
            sb.Append("</sheetData></worksheet>");
            return sb.ToString();
        }

        static string RowXml(int rowNum, string[] vals, bool header)
        {
            var sb = new StringBuilder();
            sb.Append("<row r=\"" + rowNum + "\">");
            for (int c = 0; c < vals.Length; c++)
            {
                string val = vals[c] ?? "";
                if (val.Length == 0) continue;
                string cref = ColLetter(c) + rowNum;
                sb.Append("<c r=\"" + cref + "\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
                sb.Append(XmlEsc(val));
                sb.Append("</t></is></c>");
            }
            sb.Append("</row>");
            return sb.ToString();
        }

        public static string ColLetter(int index)
        {
            string s = ""; index++;
            while (index > 0) { int m = (index - 1) % 26; s = (char)('A' + m) + s; index = (index - 1) / 26; }
            return s;
        }

        static string XmlEsc(string s)
        {
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                    .Replace("\"", "&quot;").Replace("'", "&apos;");
        }
    }

    // ============ 極簡 XLSX 讀取（支援 inlineStr / sharedStrings / str / 數字）============
    static class XlsxReader
    {
        const long MaxXmlCharacters = 64L * 1024 * 1024;

        static XmlDocument LoadSafeXml(Stream stream)
        {
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaxXmlCharacters,
                MaxCharactersFromEntities = 0
            };
            var doc = new XmlDocument { XmlResolver = null };
            using (var reader = XmlReader.Create(stream, settings)) doc.Load(reader);
            return doc;
        }

        // 回傳：每個工作表 → 列(rows) → 欄(cells)
        public static List<List<string[]>> Read(string path)
        {
            var result = new List<List<string[]>>();
            using (var zip = ZipFile.OpenRead(path))
            {
                var shared = ReadSharedStrings(zip);
                // 依 workbook.xml 的順序取得各 sheet 對應的檔名
                var sheetTargets = ReadSheetOrder(zip);
                foreach (var target in sheetTargets)
                {
                    var entry = zip.GetEntry("xl/" + target);
                    if (entry == null) continue;
                    result.Add(ReadSheet(entry, shared));
                }
                if (result.Count == 0) // 後援：直接讀所有 worksheet
                {
                    foreach (var e in zip.Entries)
                        if (e.FullName.StartsWith("xl/worksheets/") && e.FullName.EndsWith(".xml"))
                            result.Add(ReadSheet(e, shared));
                }
            }
            return result;
        }

        static List<string> ReadSharedStrings(ZipArchive zip)
        {
            var list = new List<string>();
            var e = zip.GetEntry("xl/sharedStrings.xml");
            if (e == null) return list;
            using (var st = e.Open())
            {
                var doc = LoadSafeXml(st);
                foreach (XmlNode si in doc.GetElementsByTagName("si"))
                {
                    // si 可能含多個 t（rich text）→ 串接
                    var sb = new StringBuilder();
                    foreach (XmlNode t in ((XmlElement)si).GetElementsByTagName("t")) sb.Append(t.InnerText);
                    list.Add(sb.ToString());
                }
            }
            return list;
        }

        static List<string> ReadSheetOrder(ZipArchive zip)
        {
            var targets = new List<string>();
            var wb = zip.GetEntry("xl/workbook.xml");
            var rels = zip.GetEntry("xl/_rels/workbook.xml.rels");
            if (wb == null || rels == null) return targets;

            var ridToTarget = new Dictionary<string, string>();
            using (var st = rels.Open())
            {
                var doc = LoadSafeXml(st);
                foreach (XmlNode rel in doc.GetElementsByTagName("Relationship"))
                {
                    var id = rel.Attributes["Id"]; var tg = rel.Attributes["Target"];
                    if (id != null && tg != null) ridToTarget[id.Value] = tg.Value.Replace("\\", "/");
                }
            }
            using (var st = wb.Open())
            {
                var doc = LoadSafeXml(st);
                foreach (XmlNode sheet in doc.GetElementsByTagName("sheet"))
                {
                    string rid = null;
                    foreach (XmlAttribute a in sheet.Attributes) if (a.LocalName == "id") rid = a.Value;
                    if (rid != null && ridToTarget.ContainsKey(rid)) targets.Add(ridToTarget[rid]);
                }
            }
            return targets;
        }

        static List<string[]> ReadSheet(ZipArchiveEntry entry, List<string> shared)
        {
            var rows = new List<string[]>();
            using (var st = entry.Open())
            {
                var doc = LoadSafeXml(st);
                foreach (XmlNode row in doc.GetElementsByTagName("row"))
                {
                    var cells = new Dictionary<int, string>();
                    int maxc = -1;
                    foreach (XmlNode c in ((XmlElement)row).GetElementsByTagName("c"))
                    {
                        var rAttr = c.Attributes["r"];
                        int col = rAttr != null ? RefToCol(rAttr.Value) : cells.Count;
                        string type = c.Attributes["t"] != null ? c.Attributes["t"].Value : "";
                        string val = "";
                        if (type == "inlineStr")
                        {
                            var isNode = ((XmlElement)c).GetElementsByTagName("t");
                            var sb = new StringBuilder();
                            foreach (XmlNode t in isNode) sb.Append(t.InnerText);
                            val = sb.ToString();
                        }
                        else
                        {
                            var vNode = ((XmlElement)c).GetElementsByTagName("v");
                            string raw = vNode.Count > 0 ? vNode[0].InnerText : "";
                            if (type == "s")
                            {
                                int idx; if (int.TryParse(raw, out idx) && idx >= 0 && idx < shared.Count) val = shared[idx];
                            }
                            else val = raw; // str / 數字 / 布林：原值
                        }
                        cells[col] = val;
                        if (col > maxc) maxc = col;
                    }
                    var arr = new string[maxc + 1];
                    for (int i = 0; i <= maxc; i++) arr[i] = cells.ContainsKey(i) ? cells[i] : "";
                    rows.Add(arr);
                }
            }
            return rows;
        }

        static int RefToCol(string cellRef)
        {
            int col = 0;
            foreach (char ch in cellRef)
            {
                if (ch >= 'A' && ch <= 'Z') col = col * 26 + (ch - 'A' + 1);
                else if (ch >= 'a' && ch <= 'z') col = col * 26 + (ch - 'a' + 1);
                else break;
            }
            return col - 1;
        }
    }
}
