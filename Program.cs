//. Program.cs 
//. 
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

internal static class Program
{
    // 画面描画で使う改行は CRLF に統一（cmd.exe で崩れにくい）
    private const string NL = "\r\n";

    // ===================== Windows VT(ANSI) enable =====================
    // ENABLE_VIRTUAL_TERMINAL_PROCESSING で VT(ANSI) シーケンスが解釈される。[2](https://learn.microsoft.com/en-us/windows/console/console-virtual-terminal-sequences)[3](https://github.com/MicrosoftDocs/Console-Docs/blob/main/docs/console-virtual-terminal-sequences.md)
    // DISABLE_NEWLINE_AUTO_RETURN は LF が CR を伴わない挙動になり得て画面が乱れやすいので今回は使わない。[1](https://github.com/microsoft/terminal/issues/4126)[2](https://learn.microsoft.com/en-us/windows/console/console-virtual-terminal-sequences)
    private static class WinConsole
    {
        private const int STD_OUTPUT_HANDLE = -11;
        private const int ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004;

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetStdHandle(int nStdHandle);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetConsoleMode(IntPtr hConsoleHandle, out int lpMode);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetConsoleMode(IntPtr hConsoleHandle, int dwMode);

        public static void TryEnableVT()
        {
            var h = GetStdHandle(STD_OUTPUT_HANDLE);
            if (h == IntPtr.Zero) return;
            if (!GetConsoleMode(h, out var mode)) return;

            mode |= ENABLE_VIRTUAL_TERMINAL_PROCESSING;
            SetConsoleMode(h, mode);
        }
    }

    // ===================== ANSI helpers =====================
    private static class Ansi
    {
        public const string Esc = "\u001b";
        public static string Home => $"{Esc}[H";
        public static string Clear => $"{Esc}[2J";
        public static string HideCursor => $"{Esc}[?25l";
        public static string ShowCursor => $"{Esc}[?25h";
        public static string Reset => $"{Esc}[0m";
        public static string Invert => $"{Esc}[7m";
        public static string Dim => $"{Esc}[2m";
        public static string Bold => $"{Esc}[1m";
        public static string Gray => $"{Esc}[90m";
        public static string BgBlue => $"{Esc}[44m";
        public static string BgGray => $"{Esc}[100m";
        public static string BgSel => $"{Esc}[48;5;238m"; // 256-color 環境向け（無理なら無視される）
        public static string Move(int row, int col) => $"{Esc}[{row};{col}H";
        public static string ClearLine => $"{Esc}[2K";
    }

    // ===================== Modes =====================
    private enum Mode
    {
        Normal,
        Visual,
        Command,
        Insert
    }

    // ===================== Spreadsheet model =====================
    private sealed class Cell
    {
        public string Raw = ""; // 単一行前提
    }

    private sealed class Sheet
    {
        private readonly Dictionary<(int r, int c), Cell> _cells = new();

        public int Rows { get; }
        public int Cols { get; }
        public int[] ColWidths { get; }

        public Sheet(int rows, int cols, int defaultWidth)
        {
            Rows = rows;
            Cols = cols;
            ColWidths = new int[cols];
            for (int i = 0; i < cols; i++) ColWidths[i] = defaultWidth;
        }

        public Cell GetCell(int r, int c)
        {
            if (!_cells.TryGetValue((r, c), out var cell))
            {
                cell = new Cell();
                _cells[(r, c)] = cell;
            }
            return cell;
        }

        public bool TryGetCell(int r, int c, out Cell? cell) => _cells.TryGetValue((r, c), out cell);

        public void ClearAll() => _cells.Clear();

        public void SetColWidth(int col, int width)
        {
            if (col < 0 || col >= Cols) return;
            ColWidths[col] = Math.Clamp(width, 3, 80);
        }

        public void AutoFitCol(int col, int viewTop, int viewRows)
        {
            if (col < 0 || col >= Cols) return;
            int max = 3;
            int end = Math.Min(Rows, viewTop + viewRows);

            for (int r = viewTop; r < end; r++)
            {
                if (TryGetCell(r, col, out var cell) && cell != null && !string.IsNullOrEmpty(cell.Raw))
                {
                    max = Math.Max(max, Math.Min(80, cell.Raw.Length + 1));
                }
            }
            SetColWidth(col, max);
        }

        // ---- CSV (単一行セル前提) ----
        public void SaveCsv(string path)
        {
            using var sw = new StreamWriter(path, false, new UTF8Encoding(false));
            for (int r = 0; r < Rows; r++)
            {
                for (int c = 0; c < Cols; c++)
                {
                    string raw = (TryGetCell(r, c, out var cell) && cell != null) ? (cell.Raw ?? "") : "";
                    raw = raw.Replace("\r", "").Replace("\n", ""); // 単一行保証
                    sw.Write(CsvEscape(raw));
                    if (c != Cols - 1) sw.Write(',');
                }
                sw.Write(NL); // CRLF 固定
            }
        }

        public void LoadCsv(string path)
        {
            ClearAll();
            using var sr = new StreamReader(path, new UTF8Encoding(false), true);

            int r = 0;
            while (!sr.EndOfStream && r < Rows)
            {
                var line = sr.ReadLine() ?? "";
                var cols = CsvParseLine(line);
                for (int c = 0; c < Math.Min(Cols, cols.Count); c++)
                {
                    var v = cols[c].Replace("\r", "").Replace("\n", "");
                    if (!string.IsNullOrEmpty(v))
                        GetCell(r, c).Raw = v;
                }
                r++;
            }
        }

        private static string CsvEscape(string s)
        {
            if (s.Contains(',') || s.Contains('"'))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private static List<string> CsvParseLine(string line)
        {
            var res = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];
                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            sb.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else sb.Append(ch);
                }
                else
                {
                    if (ch == '"') inQuotes = true;
                    else if (ch == ',')
                    {
                        res.Add(sb.ToString());
                        sb.Clear();
                    }
                    else sb.Append(ch);
                }
            }
            res.Add(sb.ToString());
            return res;
        }
    }

    // ===================== Formula engine (optional but kept) =====================
    private sealed class Evaluator
    {
        private readonly Sheet _sheet;
        public Evaluator(Sheet sheet) => _sheet = sheet;

        public EvalResult EvalCell(int r, int c)
        {
            var visited = new HashSet<(int r, int c)>();
            return EvalCellInternal(r, c, visited);
        }

        private EvalResult EvalCellInternal(int r, int c, HashSet<(int r, int c)> visited)
        {
            if (!visited.Add((r, c)))
                return EvalResult.Error("#CYCLE");

            if (!_sheet.TryGetCell(r, c, out var cell) || cell == null || string.IsNullOrEmpty(cell.Raw))
                return EvalResult.Number(0);

            var raw = cell.Raw.Trim();

            if (!raw.StartsWith("=", StringComparison.Ordinal))
            {
                if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var num))
                    return EvalResult.Number(num);
                return EvalResult.FromText(raw);
            }

            var expr = raw.Substring(1);
            try
            {
                var parser = new ExprParser(expr,
                    (refStr) =>
                    {
                        if (TryParseA1(refStr, out var rr, out var cc))
                            return EvalCellInternal(rr, cc, visited).AsNumberOrZero();
                        return 0;
                    },
                    (funcName, args) =>
                    {
                        var name = funcName.ToUpperInvariant();
                        var values = ExpandArgs(args, visited);
                        if (values.Count == 0) return 0;

                        return name switch
                        {
                            "SUM" => values.Sum(),
                            "AVG" => values.Average(),
                            "MIN" => values.Min(),
                            "MAX" => values.Max(),
                            _ => throw new InvalidOperationException($"Unknown func: {funcName}")
                        };
                    });

                var v = parser.ParseExpression();
                parser.ExpectEnd();
                return EvalResult.Number(v);
            }
            catch
            {
                return EvalResult.Error("#ERR");
            }
            finally
            {
                visited.Remove((r, c));
            }
        }

        private List<double> ExpandArgs(List<string> args, HashSet<(int r, int c)> visited)
        {
            var values = new List<double>();
            foreach (var a in args)
            {
                var s = a.Trim();
                if (s.Contains(':'))
                {
                    var parts = s.Split(':', 2);
                    if (TryParseA1(parts[0].Trim(), out var r1, out var c1) &&
                        TryParseA1(parts[1].Trim(), out var r2, out var c2))
                    {
                        int rr1 = Math.Min(r1, r2), rr2 = Math.Max(r1, r2);
                        int cc1 = Math.Min(c1, c2), cc2 = Math.Max(c1, c2);
                        for (int rr = rr1; rr <= rr2; rr++)
                            for (int cc = cc1; cc <= cc2; cc++)
                                values.Add(EvalCellInternal(rr, cc, visited).AsNumberOrZero());
                        continue;
                    }
                }

                if (TryParseA1(s, out var rrA, out var ccA))
                {
                    values.Add(EvalCellInternal(rrA, ccA, visited).AsNumberOrZero());
                    continue;
                }

                if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var num))
                {
                    values.Add(num);
                    continue;
                }

                values.Add(0);
            }
            return values;
        }

        public static bool TryParseA1(string s, out int row0, out int col0)
        {
            row0 = col0 = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;

            int i = 0;
            int col = 0;
            bool hasCol = false;

            while (i < s.Length && char.IsLetter(s[i]))
            {
                hasCol = true;
                char ch = char.ToUpperInvariant(s[i]);
                if (ch < 'A' || ch > 'Z') return false;
                col = col * 26 + (ch - 'A' + 1);
                i++;
            }
            if (!hasCol) return false;
            if (i >= s.Length) return false;

            int row = 0;
            bool hasRow = false;
            while (i < s.Length && char.IsDigit(s[i]))
            {
                hasRow = true;
                row = row * 10 + (s[i] - '0');
                i++;
            }
            if (!hasRow) return false;
            if (i != s.Length) return false;

            col0 = col - 1;
            row0 = row - 1;
            return row0 >= 0 && col0 >= 0;
        }
    }

    // CS0102 回避：フィールド名とメソッド名が衝突しないようにする（Str + FromText）
    private readonly struct EvalResult
    {
        public readonly string Kind; // "num" "text" "err"
        public readonly double Num;
        public readonly string Str;

        private EvalResult(string kind, double num, string str)
        {
            Kind = kind;
            Num = num;
            Str = str;
        }

        public static EvalResult Number(double v) => new("num", v, "");
        public static EvalResult FromText(string s) => new("text", 0, s);
        public static EvalResult Error(string s) => new("err", 0, s);

        public double AsNumberOrZero() => Kind == "num" ? Num : 0;

        public string ToDisplay(int maxLen)
        {
            string s = Kind switch
            {
                "num" => Num.ToString("0.########", CultureInfo.InvariantCulture),
                "text" => Str,
                "err" => Str,
                _ => ""
            };
            if (s.Length > maxLen) s = s.Substring(0, Math.Max(0, maxLen - 1)) + "…";
            return s;
        }
    }

    private sealed class ExprParser
    {
        private readonly string _s;
        private int _pos;
        private readonly Func<string, double> _cellResolver;
        private readonly Func<string, List<string>, double> _funcResolver;

        public ExprParser(string s, Func<string, double> cellResolver, Func<string, List<string>, double> funcResolver)
        { _s = s; _cellResolver = cellResolver; _funcResolver = funcResolver; }

        public double ParseExpression()
        {
            var v = ParseTerm();
            while (true)
            {
                SkipWs();
                if (Match('+')) v += ParseTerm();
                else if (Match('-')) v -= ParseTerm();
                else break;
            }
            return v;
        }

        private double ParseTerm()
        {
            var v = ParseFactor();
            while (true)
            {
                SkipWs();
                if (Match('*')) v *= ParseFactor();
                else if (Match('/')) v /= ParseFactor();
                else break;
            }
            return v;
        }

        private double ParseFactor()
        {
            SkipWs();
            if (Match('+')) return ParseFactor();
            if (Match('-')) return -ParseFactor();

            if (Match('('))
            {
                var v = ParseExpression();
                SkipWs();
                Expect(')');
                return v;
            }

            if (PeekLetter())
            {
                var ident = ReadIdentOrCell();
                SkipWs();
                if (Match('('))
                {
                    var args = ReadArgsUntil(')');
                    Expect(')');
                    return _funcResolver(ident, args);
                }
                return _cellResolver(ident);
            }

            return ReadNumber();
        }

        private List<string> ReadArgsUntil(char endCh)
        {
            var args = new List<string>();
            var sb = new StringBuilder();
            int depth = 0;

            while (_pos < _s.Length)
            {
                char ch = _s[_pos];
                if (ch == endCh && depth == 0) break;
                if (ch == '(') depth++;
                if (ch == ')') depth--;

                if (ch == ',' && depth == 0)
                {
                    args.Add(sb.ToString());
                    sb.Clear();
                    _pos++;
                    continue;
                }
                sb.Append(ch);
                _pos++;
            }

            var last = sb.ToString();
            if (!string.IsNullOrWhiteSpace(last)) args.Add(last);
            return args;
        }

        private string ReadIdentOrCell()
        {
            int start = _pos;
            while (_pos < _s.Length && char.IsLetter(_s[_pos])) _pos++;
            while (_pos < _s.Length && char.IsDigit(_s[_pos])) _pos++;
            return _s.Substring(start, _pos - start);
        }

        private double ReadNumber()
        {
            SkipWs();
            int start = _pos;
            bool dot = false;
            while (_pos < _s.Length)
            {
                char ch = _s[_pos];
                if (char.IsDigit(ch)) { _pos++; continue; }
                if (ch == '.' && !dot) { dot = true; _pos++; continue; }
                break;
            }
            if (start == _pos) throw new FormatException("number expected");
            var token = _s.Substring(start, _pos - start);
            return double.Parse(token, CultureInfo.InvariantCulture);
        }

        private void SkipWs()
        { while (_pos < _s.Length && char.IsWhiteSpace(_s[_pos])) _pos++; }

        private bool Match(char ch)
        { if (_pos < _s.Length && _s[_pos] == ch) { _pos++; return true; } return false; }

        private void Expect(char ch)
        { if (!Match(ch)) throw new FormatException($"expected '{ch}'"); }

        public void ExpectEnd()
        { SkipWs(); if (_pos != _s.Length) throw new FormatException("trailing chars"); }

        private bool PeekLetter() => _pos < _s.Length && char.IsLetter(_s[_pos]);
    }

    // ===================== UI state =====================
    private static Mode _mode = Mode.Normal;

    private static int _cursorR = 0;
    private static int _cursorC = 0;
    private static int _viewTop = 0;
    private static int _viewLeft = 0;

    private static string _status = "";
    private static string _filePath = "sheet.csv";

    // Visual selection
    private static bool _hasSelection = false;
    private static int _selAnchorR, _selAnchorC;

    // Internal clipboard (rectangular)
    private static string[,]? _clip;
    private static int _clipRows, _clipCols;

    // key sequence (gg)
    private static bool _awaitingG = false;

    // Layout
    private const int RowHeaderW = 6;
    private const int ColHeaderH = 2;
    private const int StatusH = 2;

    public static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        Console.InputEncoding = Encoding.UTF8;

        WinConsole.TryEnableVT();

        Console.Write(Ansi.HideCursor);
        string? initialPath = null;
        if (args != null && args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]))
        {
            initialPath = args[0].Trim('"');
        }
        try { Run(initialPath); }
        finally { Console.Write(Ansi.Reset + Ansi.ShowCursor); }
    }

    private static void Run(string? initialPath)
    {
        var sheet = new Sheet(rows: 500, cols: 100, defaultWidth: 10);
        var eval = new Evaluator(sheet);

        if (!string.IsNullOrWhiteSpace(initialPath))
        {
            _filePath = initialPath;
            try
            {
                if (File.Exists(initialPath))
                {
                    sheet.LoadCsv(initialPath);
                    _status = $"Loaded: {initialPath}";
                }
                else
                {
                    _status = $"File not found: {initialPath}";
                }
            }
            catch (Exception ex)
            {
                _status = $"Load error: {ex.Message}";
            }
        }

        bool running = true;
        while (running)
        {
            EnsureCursorVisible(sheet);
            Render(sheet, eval);

            var key = Console.ReadKey(intercept: true);

            switch (_mode)
            {
                case Mode.Normal:
                    running = HandleNormal(key, sheet, eval);
                    break;
                case Mode.Visual:
                    running = HandleVisual(key, sheet, eval);
                    break;
                case Mode.Command:
                    running = HandleCommand(sheet);
                    break;
                case Mode.Insert:
                    // Insert は EditCellSingleLine 内で完結
                    _mode = Mode.Normal;
                    break;
            }
        }

        Console.Write(Ansi.Clear + Ansi.Home);
    }

    // ===================== Normal mode =====================
    private static bool HandleNormal(ConsoleKeyInfo key, Sheet sheet, Evaluator eval)
    {
        _hasSelection = false;

        if (key.Key == ConsoleKey.Escape)
        {
            _awaitingG = false;
            _status = "";
            return true;
        }

        if (key.KeyChar == 'g' && key.Modifiers == 0)
        {
            if (_awaitingG)
            {
                _cursorR = 0;
                _awaitingG = false;
            }
            else _awaitingG = true;
            return true;
        }
        _awaitingG = false;

        if (( key.KeyChar == ':' || key.KeyChar == '/' ) && key.Modifiers == 0)
        {
            _mode = Mode.Command;
            return true;
        }

        if (key.KeyChar == '?' && key.Modifiers == 0)
        {
            _status = "NORMAL: hjkl move | i edit | v visual | y yank | p paste | : cmd | gg/G | 0/$";
            return true;
        }

        if (key.KeyChar == 'v' && key.Modifiers == 0)
        {
            _mode = Mode.Visual;
            _hasSelection = true;
            _selAnchorR = _cursorR;
            _selAnchorC = _cursorC;
            _status = "-- VISUAL -- (move then y to yank, Esc to cancel)";
            return true;
        }

        if (key.KeyChar == 'i' && key.Modifiers == 0)
        {
            EditCellSingleLine(sheet, eval);
            return true;
        }
        
        /*
        if (key.KeyChar == 'w' && key.Modifiers == 0)
        {
            EditCellSingleLine(sheet, eval, true);
            return true;
        }
        */

        if (key.KeyChar == 'y' && key.Modifiers == 0)
        {
            CopyRectangle(sheet, _cursorR, _cursorC, _cursorR, _cursorC);
            _status = "Yanked 1x1";
            return true;
        }

        if (key.KeyChar == 'p' && key.Modifiers == 0)
        {
            PasteAt(sheet, _cursorR, _cursorC);
            _status = _clip == null ? "Clipboard empty" : $"Pasted {_clipRows}x{_clipCols}";
            return true;
        }

        if (key.KeyChar == 'G' && key.Modifiers == 0) { _cursorR = sheet.Rows - 1; return true; }
        if (key.KeyChar == '0' && key.Modifiers == 0) { _cursorC = 0; return true; }
        if (key.KeyChar == '$' && key.Modifiers == 0) { _cursorC = sheet.Cols - 1; return true; }

        switch (key.Key)
        {
            case ConsoleKey.LeftArrow: MoveCursor(sheet, 0, -1); return true;
            case ConsoleKey.RightArrow: MoveCursor(sheet, 0, +1); return true;
            case ConsoleKey.UpArrow: MoveCursor(sheet, -1, 0); return true;
            case ConsoleKey.DownArrow: MoveCursor(sheet, +1, 0); return true;
            case ConsoleKey.PageUp: _cursorR = Math.Max(0, _cursorR - VisibleRowCount()); return true;
            case ConsoleKey.PageDown: _cursorR = Math.Min(sheet.Rows - 1, _cursorR + VisibleRowCount()); return true;
        }

        if (key.Modifiers == 0)
        {
            switch (key.KeyChar)
            {
                case 'h': MoveCursor(sheet, 0, -1); return true;
                case 'l': MoveCursor(sheet, 0, +1); return true;
                case 'k': MoveCursor(sheet, -1, 0); return true;
                case 'j': MoveCursor(sheet, +1, 0); return true;
            }
        }

        return true;
    }

    // ===================== Visual mode (rect selection) =====================
    private static bool HandleVisual(ConsoleKeyInfo key, Sheet sheet, Evaluator eval)
    {
        _hasSelection = true;

        if (key.Key == ConsoleKey.Escape)
        {
            _mode = Mode.Normal;
            _hasSelection = false;
            _status = "";
            return true;
        }

        if (key.KeyChar == 'y' && key.Modifiers == 0)
        {
            var (r1, c1, r2, c2) = GetSelRect();
            CopyRectangle(sheet, r1, c1, r2, c2);
            _status = $"Yanked {Math.Abs(r2 - r1) + 1}x{Math.Abs(c2 - c1) + 1}";
            _mode = Mode.Normal;
            _hasSelection = false;
            return true;
        }

        switch (key.Key)
        {
            case ConsoleKey.LeftArrow: MoveCursor(sheet, 0, -1); return true;
            case ConsoleKey.RightArrow: MoveCursor(sheet, 0, +1); return true;
            case ConsoleKey.UpArrow: MoveCursor(sheet, -1, 0); return true;
            case ConsoleKey.DownArrow: MoveCursor(sheet, +1, 0); return true;
        }
        if (key.Modifiers == 0)
        {
            switch (key.KeyChar)
            {
                case 'h': MoveCursor(sheet, 0, -1); return true;
                case 'l': MoveCursor(sheet, 0, +1); return true;
                case 'k': MoveCursor(sheet, -1, 0); return true;
                case 'j': MoveCursor(sheet, +1, 0); return true;
            }
        }

        return true;
    }

    // ===================== Command mode =====================
    private static bool HandleCommand(Sheet sheet)
    {
        string? cmd = ReadLineAtBottom(":", "");  //. この ":" がデフォルト状態を指示している？
        _mode = Mode.Normal;

        if (cmd == null) { _status = ""; return true; }
        cmd = cmd.Trim();
        if (cmd.Length == 0) { _status = ""; return true; }

        var parts = SplitArgs(cmd);
        var op = parts[0];

        try
        {
            if (op is "q" or "quit") return false;

            if (op is "w" or "write")
            {
                string path = (parts.Count >= 2) ? parts[1] : _filePath;
                _filePath = path;
                sheet.SaveCsv(path);
                _status = $"Written: {path}";
                return true;
            }

            if (op is "e" or "edit")
            {
                if (parts.Count < 2) { _status = "Usage: :e <file.csv>"; return true; }
                string path = parts[1];
                _filePath = path;
                sheet.LoadCsv(path);
                _status = $"Loaded: {path}";
                return true;
            }

            if (op == "set")
            {
                if (parts.Count >= 2 && parts[1] == "width")
                {
                    if (parts.Count == 3)
                    {
                        int w = int.Parse(parts[2], CultureInfo.InvariantCulture);
                        sheet.SetColWidth(_cursorC, w);
                        _status = $"Col {ColName(_cursorC)} width={w}";
                        return true;
                    }
                    if (parts.Count == 4)
                    {
                        int col = ParseCol(parts[2]);
                        int w = int.Parse(parts[3], CultureInfo.InvariantCulture);
                        sheet.SetColWidth(col, w);
                        _status = $"Col {ColName(col)} width={w}";
                        return true;
                    }
                    _status = "Usage: :set width <n>  OR  :set width <Col> <n>";
                    return true;
                }

                if (parts.Count >= 2 && parts[1] == "auto")
                {
                    sheet.AutoFitCol(_cursorC, _viewTop, VisibleRowCount());
                    _status = $"AutoFit {ColName(_cursorC)} width={sheet.ColWidths[_cursorC]}";
                    return true;
                }

                _status = "Commands: :set width ... | :set auto";
                return true;
            }

            if (op is "help" or "h")
            {
                _status = "CMD: :w [file] :e <file> :q :set width [Col] <n> :set auto";
                return true;
            }

            _status = $"Unknown command: {op} (try :help)";
            return true;
        }
        catch (Exception ex)
        {
            _status = $"Command error: {ex.Message}";
            return true;
        }
    }

    private static List<string> SplitArgs(string s)
    {
        var res = new List<string>();
        var sb = new StringBuilder();
        bool inQ = false;
        for (int i = 0; i < s.Length; i++)
        {
            char ch = s[i];
            if (inQ)
            {
                if (ch == '"') inQ = false;
                else sb.Append(ch);
            }
            else
            {
                if (char.IsWhiteSpace(ch))
                {
                    if (sb.Length > 0) { res.Add(sb.ToString()); sb.Clear(); }
                }
                else if (ch == '"') inQ = true;
                else sb.Append(ch);
            }
        }
        if (sb.Length > 0) res.Add(sb.ToString());
        return res;
    }

    private static int ParseCol(string colName)
    {
        colName = colName.Trim().ToUpperInvariant();
        int col = 0;
        foreach (var ch in colName)
        {
            if (ch < 'A' || ch > 'Z') throw new FormatException("Bad column");
            col = col * 26 + (ch - 'A' + 1);
        }
        return col - 1;
    }

    // ===================== Single-line cell editor =====================
    private static void EditCellSingleLine(Sheet sheet, Evaluator eval)
    {
        _mode = Mode.Insert;

        var cell = sheet.GetCell(_cursorR, _cursorC);
        string original = cell.Raw ?? "";

        string prompt = $"Edit {ColName(_cursorC)}{_cursorR + 1}> ";
        string? input = ReadLineAtBottom(prompt, original);

        if (input == null)
        {
            cell.Raw = original;
            _status = "Edit cancelled";
        }
        else
        {
            input = input.Replace("\r", "").Replace("\n", ""); // 単一行保証
            cell.Raw = input;
            var r = eval.EvalCell(_cursorR, _cursorC);
            _status = $"Saved => {r.ToDisplay(30)}";
        }

        _mode = Mode.Normal;
    }

    // ===================== Clipboard (rectangular copy/paste) =====================
    private static void CopyRectangle(Sheet sheet, int r1, int c1, int r2, int c2)
    {
        int top = Math.Min(r1, r2), bot = Math.Max(r1, r2);
        int left = Math.Min(c1, c2), right = Math.Max(c1, c2);

        _clipRows = bot - top + 1;
        _clipCols = right - left + 1;
        _clip = new string[_clipRows, _clipCols];

        for (int r = 0; r < _clipRows; r++)
        {
            for (int c = 0; c < _clipCols; c++)
            {
                int rr = top + r;
                int cc = left + c;
                string raw = (sheet.TryGetCell(rr, cc, out var cell) && cell != null) ? (cell.Raw ?? "") : "";
                _clip[r, c] = raw;
            }
        }
    }

    private static void PasteAt(Sheet sheet, int targetR, int targetC)
    {
        if (_clip == null) return;
        for (int r = 0; r < _clipRows; r++)
        {
            for (int c = 0; c < _clipCols; c++)
            {
                int rr = targetR + r;
                int cc = targetC + c;
                if (rr < 0 || rr >= sheet.Rows || cc < 0 || cc >= sheet.Cols) continue;
                sheet.GetCell(rr, cc).Raw = _clip[r, c] ?? "";
            }
        }
    }

    private static (int r1, int c1, int r2, int c2) GetSelRect()
        => (_selAnchorR, _selAnchorC, _cursorR, _cursorC);

    // ===================== Rendering & scrolling =====================
    private static void MoveCursor(Sheet sheet, int dr, int dc)
    {
        _cursorR = Math.Clamp(_cursorR + dr, 0, sheet.Rows - 1);
        _cursorC = Math.Clamp(_cursorC + dc, 0, sheet.Cols - 1);
    }

    private static int VisibleRowCount()
    {
        int h = SafeWindowHeight();
        int gridH = Math.Max(5, h - ColHeaderH - StatusH);
        return gridH;
    }

    private static int SafeWindowWidth()
    {
        try { return Console.WindowWidth; } catch { return 120; }
    }
    private static int SafeWindowHeight()
    {
        try { return Console.WindowHeight; } catch { return 40; }
    }

    private static void EnsureCursorVisible(Sheet sheet)
    {
        int w = SafeWindowWidth();
        int h = SafeWindowHeight();

        int gridW = Math.Max(20, w - RowHeaderW);
        int gridH = Math.Max(5, h - ColHeaderH - StatusH);

        int rowsVisible = gridH;
        if (_cursorR < _viewTop) _viewTop = _cursorR;
        if (_cursorR >= _viewTop + rowsVisible) _viewTop = _cursorR - rowsVisible + 1;
        if (_viewTop < 0) _viewTop = 0;

        if (_cursorC < _viewLeft) _viewLeft = _cursorC;

        while (!IsColVisible(sheet, _cursorC, _viewLeft, gridW))
        {
            _viewLeft++;
            if (_viewLeft >= sheet.Cols) { _viewLeft = sheet.Cols - 1; break; }
        }

        if (_viewLeft < 0) _viewLeft = 0;
    }

    private static bool IsColVisible(Sheet sheet, int col, int viewLeft, int gridW)
    {
        int x = 0;
        for (int c = viewLeft; c < sheet.Cols; c++)
        {
            int cw = sheet.ColWidths[c];
            if (c == col)
                return (x < gridW);

            x += cw;
            if (x >= gridW) break;
        }
        return false;
    }

    private static void Render(Sheet sheet, Evaluator eval)
    {
        int w = SafeWindowWidth();
        int h = SafeWindowHeight();

        int gridW = Math.Max(20, w - RowHeaderW);
        int gridH = Math.Max(5, h - ColHeaderH - StatusH);

        // visible columns by width
        var visCols = new List<int>();
        int used = 0;
        for (int c = _viewLeft; c < sheet.Cols; c++)
        {
            int cw = sheet.ColWidths[c];
            if (used + cw > gridW) break;
            visCols.Add(c);
            used += cw;
        }
        if (visCols.Count == 0) visCols.Add(_viewLeft);

        int rowsVisible = gridH;

        // selection bounds
        bool selActive = (_mode == Mode.Visual) && _hasSelection;
        int selTop = 0, selBot = -1, selLeft = 0, selRight = -1;
        if (selActive)
        {
            var (r1, c1, r2, c2) = GetSelRect();
            selTop = Math.Min(r1, r2);
            selBot = Math.Max(r1, r2);
            selLeft = Math.Min(c1, c2);
            selRight = Math.Max(c1, c2);
        }

        var sb = new StringBuilder(Ansi.Clear.Length + 20000);
        sb.Append(Ansi.Home);
        sb.Append(Ansi.Clear);

        // header
        sb.Append(Ansi.Bold).Append(Ansi.BgGray);
        sb.Append("".PadRight(RowHeaderW));
        foreach (var c in visCols)
        {
            sb.Append(Fit(ColName(c), sheet.ColWidths[c]));
        }
        sb.Append(Ansi.Reset).Append(NL);

        // divider
        sb.Append(Ansi.Gray);
        sb.Append(new string('-', RowHeaderW));
        foreach (var c in visCols)
            sb.Append(new string('-', sheet.ColWidths[c]));
        sb.Append(Ansi.Reset).Append(NL);

        // rows
        for (int vr = 0; vr < rowsVisible; vr++)
        {
            int rr = _viewTop + vr;
            if (rr >= sheet.Rows) { sb.Append(NL); continue; }

            sb.Append(Ansi.Bold).Append(Ansi.Gray);
            sb.Append((rr + 1).ToString().PadLeft(RowHeaderW - 1)).Append(' ');
            sb.Append(Ansi.Reset);

            foreach (var cc in visCols)
            {
                int cw = sheet.ColWidths[cc];

                string raw = (sheet.TryGetCell(rr, cc, out var cell) && cell != null) ? (cell.Raw ?? "") : "";
                string disp;

                if (string.IsNullOrEmpty(raw))
                    disp = "";
                else if (raw.StartsWith("=", StringComparison.Ordinal))
                    disp = eval.EvalCell(rr, cc).ToDisplay(cw);
                else
                    disp = raw;

                disp = Fit(disp, cw);

                bool isCur = (rr == _cursorR && cc == _cursorC);
                bool isSel = selActive && (rr >= selTop && rr <= selBot && cc >= selLeft && cc <= selRight);

                if (isSel) sb.Append(Ansi.BgSel);
                if (isCur) sb.Append(Ansi.Invert);

                sb.Append(disp);

                if (isCur || isSel) sb.Append(Ansi.Reset);
            }

            sb.Append(NL);
        }

        // status
        int statusTop = ColHeaderH + rowsVisible;

        sb.Append(Ansi.Move(statusTop + 1, 1));
        sb.Append(Ansi.ClearLine);
        sb.Append(Ansi.BgBlue).Append(Ansi.Bold);

        string addr = $"{ColName(_cursorC)}{_cursorR + 1}";
        string modeTxt = _mode switch
        {
            Mode.Normal => "NORMAL",
            Mode.Visual => "VISUAL",
            Mode.Command => "COMMAND",
            Mode.Insert => "INSERT",
            _ => "NORMAL"
        };

        string rawCur = sheet.GetCell(_cursorR, _cursorC).Raw ?? "";
        rawCur = Trim(rawCur, 40);

        string leftInfo = $" {modeTxt}  {addr}  w={sheet.ColWidths[_cursorC]}  raw:{rawCur}";
        string rightInfo = string.IsNullOrEmpty(_status) ? "" : $"  {_status}";
        string line = (leftInfo + rightInfo).PadRight(w);
        sb.Append(line.Substring(0, Math.Min(w, line.Length)));
        sb.Append(Ansi.Reset);

        sb.Append(Ansi.Move(statusTop + 2, 1));
        sb.Append(Ansi.ClearLine);
        sb.Append(Ansi.Dim);

        string help = _mode switch
        {
            Mode.Normal => "hjkl move | i edit | v visual | y yank | p paste | : cmd | ? help",
            Mode.Visual => "VISUAL: move then y yank | Esc cancel",
            Mode.Command => "COMMAND: Enter executes | Esc cancels",
            Mode.Insert => "INSERT: Enter commit | Esc cancel",
            _ => ""
        };
        sb.Append(help.PadRight(w).Substring(0, Math.Min(w, help.Length)));
        sb.Append(Ansi.Reset);

        Console.Write(sb.ToString());
        _status = "";
    }

    private static string Fit(string s, int width)
    {
        if (width <= 0) return "";
        if (s.Length > width) return s.Substring(0, width);
        if (s.Length < width) return s.PadRight(width);
        return s;
    }

    private static string Trim(string s, int max)
    {
        if (s.Length <= max) return s;
        return s.Substring(0, Math.Max(0, max - 1)) + "…";
    }

    private static string ColName(int col0)
    {
        int n = col0 + 1;
        var sb = new StringBuilder();
        while (n > 0)
        {
            int rem = (n - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            n = (n - 1) / 26;
        }
        return sb.ToString();
    }

    // ===================== Bottom-line input (Command / Edit) =====================
    private static string? ReadLineAtBottom(string prompt, string defaultValue)
    {
        int w = SafeWindowWidth();
        int h = SafeWindowHeight();

        var sb = new StringBuilder(defaultValue ?? "");

        void Redraw()
        {
            Console.Write(Ansi.Move(h, 1) + Ansi.ClearLine + Ansi.Reset);
            string text = prompt + sb.ToString();
            if (text.Length > w - 1)
                text = text.Substring(text.Length - (w - 1));
            Console.Write(text);
            Console.Out.Flush();
        }

        Redraw();
        Console.Write(Ansi.ShowCursor);

        while (true)
        {
            var k = Console.ReadKey(true);

            if (k.Key == ConsoleKey.Escape)
            {
                Console.Write(Ansi.Move(h, 1) + Ansi.ClearLine);
                Console.Write(Ansi.HideCursor);
                return null;
            }
            if (k.Key == ConsoleKey.Enter)
            {
                Console.Write(Ansi.Move(h, 1) + Ansi.ClearLine);
                Console.Write(Ansi.HideCursor);
                return sb.ToString();
            }
            if (k.Key == ConsoleKey.Backspace)
            {
                if (sb.Length > 0)
                {
                    sb.Length--;
                    Redraw();
                }
                continue;
            }

            if (!char.IsControl(k.KeyChar))
            {
                sb.Append(k.KeyChar);
                Redraw();
            }
        }
    }
}
