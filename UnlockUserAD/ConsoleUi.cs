using Pastel;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;

namespace ADUtils
{
    /// <summary>
    /// Screen layout for the console: header, breadcrumbs, framed panels, tables and status lines.
    ///
    /// Exists because the previous output had three concrete readability problems: there was no way
    /// to tell where you were in the menu tree, the user-info dump was an unaligned wall of text,
    /// and success versus failure was signalled by colour alone -- which is lost the moment output
    /// is piped or logged.
    ///
    /// Everything renders through <see cref="AppLog"/>, so the session trace keeps a plain-text copy
    /// of exactly what the operator saw.
    ///
    /// Layout rule to preserve when editing: Pastel injects invisible ANSI escapes, so a coloured
    /// string's .Length is not its width. Always pad and measure the RAW text, then colour the
    /// finished segment -- or use <see cref="VisibleLength"/>.
    /// </summary>
    internal static class ConsoleUi
    {
        /// <summary>Inner width of panels and rules. Clamped to the window if it is narrower.</summary>
        private const int PreferredWidth = 74;

        private static readonly Regex AnsiEscape = new Regex(@"\x1B\[[0-9;]*m", RegexOptions.Compiled);

        private static bool _unicode;

        // Colour roles named by meaning rather than by colour, so usage stays consistent.
        private static readonly Color ChromeColor = Color.SteelBlue;      // frames, rules
        private static readonly Color LabelColor = Color.DarkGray;        // field labels
        private static readonly Color AccentColor = Color.MediumPurple;   // things to type
        private static readonly Color GoodColor = Color.LimeGreen;
        private static readonly Color CautionColor = Color.DarkGoldenrod;
        private static readonly Color BadColor = Color.IndianRed;
        private static readonly Color TitleColor = Color.WhiteSmoke;
        private static readonly Color MutedColor = Color.SlateGray;

        /// <summary>
        /// Switches the console to UTF-8 so box-drawing and status glyphs render.
        ///
        /// Without this, a console on a legacy code page best-fit-maps anything non-ASCII -- em
        /// dashes silently became hyphens, and box characters would come out as mojibake. If the
        /// switch fails we fall back to pure ASCII framing rather than printing rubbish.
        /// </summary>
        public static void Initialize()
        {
            try
            {
                Console.OutputEncoding = Encoding.UTF8;
                _unicode = true;
            }
            catch (Exception)
            {
                _unicode = false;      // redirected output, or a console that refuses UTF-8
            }
        }// end of Initialize

        // ---- glyphs, with an ASCII fallback ------------------------------------------------

        private static char TopLeft => _unicode ? '┌' : '+';
        private static char TopRight => _unicode ? '┐' : '+';
        private static char BottomLeft => _unicode ? '└' : '+';
        private static char BottomRight => _unicode ? '┘' : '+';
        private static char Horizontal => _unicode ? '─' : '-';
        private static char Vertical => _unicode ? '│' : '|';
        private static char HeavyRule => _unicode ? '═' : '=';
        private static string Bullet => _unicode ? "▸" : ">";
        private static string Dot => _unicode ? "●" : "*";

        private static string OkGlyph => _unicode ? "✔" : "[OK]";
        private static string WarnGlyph => _unicode ? "⚠" : "[!]";
        private static string FailGlyph => _unicode ? "✖" : "[X]";

        private static int Width
        {
            get
            {
                try
                {
                    return Math.Max(44, Math.Min(PreferredWidth, Console.WindowWidth - 4));
                }
                catch (IOException)
                {
                    return PreferredWidth;      // redirected output has no window
                }
            }
        }

        // ---- header and navigation ---------------------------------------------------------

        /// <summary>
        /// Banner and context bar: which account you are authenticated as, and against which server.
        /// Worth two lines -- this tool makes privileged changes, and "which account am I?" is the
        /// first thing worth knowing before making one.
        /// </summary>
        public static void Header()
        {
            int w = Width;
            string who = AdminSession.IsSet
                ? $"{AdminSession.Username} @ {AdminSession.Domain} {Dot} {AdminSession.DomainName}"
                : "not authenticated";
            string time = DateTime.Now.ToString("HH:mm:ss");
            string rule = new string(HeavyRule, w);

            AppLog.Screen("");
            AppLog.Screen($"  {rule}", ChromeColor);
            AppLog.Screen($"  ADUtils {Dot} Active Directory Utility", TitleColor);
            AppLog.Screen($"  {who.PadRight(Math.Max(0, w - time.Length))}{time}", LabelColor);
            AppLog.Screen($"  {rule}", ChromeColor);
        }// end of Header

        /// <summary>Shows the path through the menus, e.g. "Main > Group Management".</summary>
        public static void Breadcrumb(params string[] path)
        {
            AppLog.Screen("");
            AppLog.Screen($"  {string.Join($" {Bullet} ", path)}", MutedColor);
        }// end of Breadcrumb

        /// <summary>Renders a numbered menu; numbering is 1-based and matches the switch cases.</summary>
        public static void Menu(string title, params string[] items)
        {
            AppLog.Screen("");
            AppLog.Screen($"  {title}", TitleColor);
            AppLog.Screen("");
            for (int i = 0; i < items.Length; i++)
            {
                AppLog.Screen($"    {(i + 1).ToString().Pastel(AccentColor)}   {items[i]}");
            }
            AppLog.Screen("");
        }// end of Menu

        /// <summary>An input prompt in the house style. No trailing newline.</summary>
        public static void Prompt(string label)
        {
            AppLog.Prompt($"  {Bullet.Pastel(AccentColor)} {label}: ");
        }// end of Prompt

        /// <summary>A prompt that names 'exit' as the way out, since every loop honours it.</summary>
        public static void PromptWithExit(string label)
        {
            AppLog.Prompt($"  {Bullet.Pastel(AccentColor)} {label} (or {"exit".Pastel(AccentColor)}): ");
        }// end of PromptWithExit

        // ---- status lines -------------------------------------------------------------------

        // Status carries a glyph as well as colour, so it survives being logged or piped.
        public static void Ok(string message) => AppLog.Info($"  {OkGlyph} {message}", GoodColor);
        public static void Warn(string message) => AppLog.Warn($"  {WarnGlyph} {message}", color: CautionColor);
        public static void Fail(string message) => AppLog.Warn($"  {FailGlyph} {message}", color: BadColor);
        public static void Fail(string message, Exception ex) => AppLog.Error($"  {FailGlyph} {message}", ex, BadColor);
        public static void Note(string message) => AppLog.Screen($"    {message}", MutedColor);
        public static void Blank() => AppLog.Blank();

        // ---- panels and tables --------------------------------------------------------------

        /// <summary>
        /// A framed key/value panel. Labels pad to a common width so values line up -- the point of
        /// the exercise, since the old single-WriteLine dump did not align at all.
        /// </summary>
        public static void Panel(string title, IEnumerable<(string Label, string Value)> fields)
        {
            var rows = fields.ToList();
            int w = Width;
            int labelWidth = rows.Count == 0 ? 0 : Math.Min(22, rows.Max(r => r.Label.Length) + 2);

            // Build the title bar as raw text first, then colour it, so the dashes count correctly.
            string titleBar = $"{TopLeft}{Horizontal} {title} ";
            titleBar += new string(Horizontal, Math.Max(0, w - VisibleLength(titleBar) - 1)) + TopRight;

            AppLog.Screen("");
            AppLog.Screen($"  {titleBar}", ChromeColor);

            foreach (var (label, value) in rows)
            {
                string text = value ?? "N/A";
                string currentLabel = label;      // a foreach deconstruction variable is read-only

                // Coloured values (state indicators) carry ANSI escapes, so their length is not their
                // width; don't try to wrap those -- they are short by construction.
                var lines = ContainsAnsi(text)
                    ? new List<string> { text }
                    : Wrap(text, w - labelWidth - 5).ToList();

                foreach (string line in lines)
                {
                    // Close the right border. The value may carry ANSI escapes, so the padding has
                    // to be computed from the visible width, not from string.Length.
                    int used = 2 + labelWidth + VisibleLength(line);
                    string tail = new string(' ', Math.Max(0, w - used - 2));

                    AppLog.Screen($"  {Vertical.ToString().Pastel(ChromeColor)}"
                                + $"  {currentLabel.PadRight(labelWidth).Pastel(LabelColor)}{line}{tail}"
                                + Vertical.ToString().Pastel(ChromeColor));

                    currentLabel = string.Empty;   // only the first wrapped line carries the label
                }
            }

            AppLog.Screen($"  {BottomLeft}{new string(Horizontal, Math.Max(0, w - 2))}{BottomRight}", ChromeColor);
        }// end of Panel

        /// <summary>
        /// A column table for report output. Columns size to their widest cell so results scan
        /// cleanly, with the row count underneath so it is obvious when a report is empty.
        /// </summary>
        public static void Table(string[] headers, IEnumerable<string[]> rows)
        {
            var data = rows.ToList();
            if (data.Count == 0)
            {
                Note("(nothing to report)");
                return;
            }

            int columns = headers.Length;
            var widths = new int[columns];
            for (int c = 0; c < columns; c++)
            {
                widths[c] = headers[c].Length;
                foreach (var row in data)
                {
                    if (c < row.Length && row[c] != null) widths[c] = Math.Max(widths[c], row[c].Length);
                }
            }

            AppLog.Screen("");
            AppLog.Screen("  " + string.Join("  ", headers.Select((h, c) => h.PadRight(widths[c]))), TitleColor);
            AppLog.Screen("  " + string.Join("  ", widths.Select(x => new string(Horizontal, x))), ChromeColor);

            foreach (var row in data)
            {
                AppLog.Screen("  " + string.Join("  ", Enumerable.Range(0, columns)
                    .Select(c => (c < row.Length ? row[c] ?? string.Empty : string.Empty).PadRight(widths[c]))));
            }

            AppLog.Screen("");
            AppLog.Screen($"  {data.Count} row(s).", MutedColor);
        }// end of Table

        /// <summary>A state indicator for panel values, e.g. "● Enabled" / "● LOCKED".</summary>
        public static string State(bool good, string text)
        {
            return $"{Dot} {text}".Pastel(good ? GoodColor : BadColor);
        }// end of State

        /// <summary>Asks a yes/no question. Anything other than Y is a no.</summary>
        public static bool Confirm(string question)
        {
            AppLog.Prompt($"  {Bullet.Pastel(AccentColor)} {question} ({"Y/N".Pastel(AccentColor)}): ");
            return ConsoleInput.ReadTrimmedUpper() == "Y";
        }// end of Confirm

        // ---- helpers -------------------------------------------------------------------------

        /// <summary>Printable width, ignoring Pastel's ANSI escapes.</summary>
        private static int VisibleLength(string text)
        {
            return string.IsNullOrEmpty(text) ? 0 : AnsiEscape.Replace(text, string.Empty).Length;
        }// end of VisibleLength

        private static bool ContainsAnsi(string text)
        {
            return !string.IsNullOrEmpty(text) && AnsiEscape.IsMatch(text);
        }// end of ContainsAnsi

        /// <summary>Splits a value across lines, breaking on spaces where it can.</summary>
        private static IEnumerable<string> Wrap(string text, int width)
        {
            if (width <= 0 || text.Length <= width)
            {
                yield return text;
                yield break;
            }

            int index = 0;
            while (index < text.Length)
            {
                int take = Math.Min(width, text.Length - index);
                if (take == width && index + take < text.Length)
                {
                    int space = text.LastIndexOf(' ', index + take - 1, take);
                    if (space > index) take = space - index + 1;
                }
                yield return text.Substring(index, take).TrimEnd();
                index += take;
            }
        }// end of Wrap
    }// end of class
}// end of namespace
