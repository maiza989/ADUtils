namespace ADUtils
{
    /// <summary>
    /// Console input helpers that never return null.
    ///
    /// Console.ReadLine() returns null when stdin reaches end-of-stream -- a closed console,
    /// piped input, or a scheduled task. Every caller in this project immediately chained
    /// .Trim()/.ToLower() onto the result, so any of those produced a NullReferenceException
    /// that unwound the menu loop back to the login prompt.
    /// </summary>
    public static class ConsoleInput
    {
        /// <summary>
        /// True once stdin has reached end-of-stream. Callers driving a retry loop must check this,
        /// otherwise substituting an empty string for null turns "no more input" into an infinite
        /// loop -- e.g. the credential prompt re-asking forever under a scheduled task.
        /// </summary>
        public static bool EndOfInput { get; private set; }

        /// <summary>Reads a line, returning an empty string at end-of-stream.</summary>
        public static string ReadLine()
        {
            string line = Console.ReadLine();
            if (line == null)
            {
                EndOfInput = true;
                return string.Empty;
            }
            return line;
        }// end of ReadLine

        /// <summary>Reads a trimmed line, returning an empty string at end-of-stream.</summary>
        public static string ReadTrimmed()
        {
            return ReadLine().Trim();
        }// end of ReadTrimmed

        /// <summary>Reads a trimmed, lower-cased line -- for menu choices and 'exit' checks.</summary>
        public static string ReadTrimmedLower()
        {
            return ReadTrimmed().ToLowerInvariant();
        }// end of ReadTrimmedLower

        /// <summary>Reads a trimmed, upper-cased line -- for Y/N confirmations.</summary>
        public static string ReadTrimmedUpper()
        {
            return ReadTrimmed().ToUpperInvariant();
        }// end of ReadTrimmedUpper
    }// end of class
}// end of namespace
