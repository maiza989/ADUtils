using System;
using System.IO;

namespace UnlockUserAD
{
    public class AuditLogManager
    {
        private static readonly string logFilePath = "audit_log.txt";

        public AuditLogManager()
        {
            // Ensure the log file is created and append mode is set
            File.AppendAllText(logFilePath, $"Audit log started at {DateTime.Now}\n");
        }
        public void Log(string message)
        {
            string logEntry = $"{DateTime.Now}: {message}\n";
            File.AppendAllText(logFilePath, logEntry);
        }
        public void RedirectConsoleOutput()
        {
            FileStream fileStream = new FileStream(logFilePath, FileMode.Append, FileAccess.Write);
            StreamWriter streamWriter = new StreamWriter(fileStream);
            streamWriter.AutoFlush = true;
            Console.SetOut(streamWriter);
            Console.SetError(streamWriter); // Optional: Redirect error output as well
        }
    }// end of class
}// end of namespace
