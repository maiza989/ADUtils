using System;
using System.IO;

namespace UnlockUserAD
{
    public class AuditLogManager
    {
        private static readonly string logFilePath = @"PATH\YOU\WANT";
        
        public AuditLogManager()
        {
            // Ensure the log file is created and append mode is set
            File.AppendAllText(logFilePath, $"--------------------------------------------------------------------------\n" +
                                            $"\t\t\t\tAudit log started at {DateTime.Now}\n" +
                                            $"--------------------------------------------------------------------------\n");
        }
        public void Log(string message)
        {
            string logEntry = $"{DateTime.Now}: {message}\n";
            File.AppendAllText(logFilePath, logEntry);
        }
        public void RedirectConsoleOutput()
        {
            FileStream fileStream = new FileStream(logFilePath, FileMode.Append, FileAccess.Write);
            StreamWriter fileWriter = new StreamWriter(fileStream) { AutoFlush = true };

            TextWriter consoleWriter = Console.Out;
            TextWriter dualWriter = new DualWriterManager(consoleWriter, fileWriter);
            Console.SetOut(dualWriter);
            Console.SetError(dualWriter); // Optional: Redirect error output as well
        }
    }// end of class
}// end of namespace
