using NLog;
using Pastel;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace ADUtils
{
    /// <summary>
    /// Writes one message to two places: the console, for the operator standing in front of the
    /// tool, and NLog, for the persistent record.
    ///
    /// Centralised rather than repeating a Console.WriteLine and a Logger call at every site,
    /// because the two drifted apart constantly -- most failures were only ever printed, so the
    /// stack trace was lost the moment the window scrolled.
    ///
    /// Console text is colourised with Pastel; the logged copy has the ANSI escapes stripped so
    /// the log files stay readable in a text editor.
    /// </summary>
    internal static class AppLog
    {
        // Matches the SGR escape sequences Pastel emits, e.g. ESC[38;2;205;92;92m
        private static readonly Regex AnsiEscape = new Regex(@"\x1B\[[0-9;]*m", RegexOptions.Compiled);

        /// <summary>Operator-facing progress or result. Console + log at Info.</summary>
        public static void Info(string message, Color? color = null,
                                [CallerFilePath] string file = "", [CallerMemberName] string member = "")
        {
            Console.WriteLine(color.HasValue ? message.Pastel(color.Value) : message);
            LoggerFor(file).Info("{0}: {1}", member, Clean(message));
        }// end of Info

        /// <summary>
        /// Something the operator needs to know about but which is not a failure -- a skipped DC,
        /// a partial success, a step that has to be done by hand. Console + log at Warn.
        ///
        /// Takes the exception in the same position as <see cref="Error"/> so a call site can be
        /// switched between the two levels without reordering arguments.
        /// </summary>
        public static void Warn(string message, Exception ex = null, Color? color = null,
                                [CallerFilePath] string file = "", [CallerMemberName] string member = "")
        {
            Console.WriteLine(color.HasValue ? message.Pastel(color.Value) : message.Pastel(Color.DarkGoldenrod));

            Logger logger = LoggerFor(file);
            if (ex != null)
            {
                logger.Warn(ex, "{0}: {1}", member, Clean(message));
            }
            else
            {
                logger.Warn("{0}: {1}", member, Clean(message));
            }
        }// end of Warn

        /// <summary>
        /// A failure. Pass the exception so the log gets the full stack trace even though the
        /// console only shows the message. Console + log at Error.
        /// </summary>
        public static void Error(string message, Exception ex = null, Color? color = null,
                                 [CallerFilePath] string file = "", [CallerMemberName] string member = "")
        {
            Console.WriteLine(color.HasValue ? message.Pastel(color.Value) : message.Pastel(Color.IndianRed));

            Logger logger = LoggerFor(file);
            if (ex != null)
            {
                logger.Error(ex, "{0}: {1}", member, Clean(message));
            }
            else
            {
                logger.Error("{0}: {1}", member, Clean(message));
            }
        }// end of Error

        /// <summary>
        /// Log-only detail, with nothing printed. For context worth keeping in the file but not
        /// worth adding to the operator's screen.
        /// </summary>
        public static void Detail(string message,
                                  [CallerFilePath] string file = "", [CallerMemberName] string member = "")
        {
            LoggerFor(file).Debug("{0}: {1}", member, Clean(message));
        }// end of Detail

        /// <summary>
        /// Names the logger after the calling source file, so log lines identify their origin
        /// without every class having to declare its own logger field.
        /// </summary>
        private static Logger LoggerFor(string filePath)
        {
            string name = string.IsNullOrEmpty(filePath) ? "ADUtils" : Path.GetFileNameWithoutExtension(filePath);
            return LogManager.GetLogger(name);
        }// end of LoggerFor

        /// <summary>Strips Pastel's ANSI colour codes so log files are plain text.</summary>
        private static string Clean(string message)
        {
            return string.IsNullOrEmpty(message) ? string.Empty : AnsiEscape.Replace(message, string.Empty).Trim();
        }// end of Clean
    }// end of class
}// end of namespace
