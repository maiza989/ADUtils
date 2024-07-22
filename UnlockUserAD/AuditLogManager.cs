using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;

namespace ADUtils
{
    public class AuditLogManager
    {
        private static readonly string BaseLogDirectory = @"H:\IT\Maitham's Cave\ADUtil\Logs";                                   // Replace with your desire log location 
        private string logFilePath;

        /// <summary>
        /// A constructor that create and ensure the for log file exists.
        /// </summary>
        /// <param name="adminUsername"></param>
        public AuditLogManager(string adminUsername)
        {
            try
            {
            logFilePath = Path.Combine(BaseLogDirectory, $"{adminUsername}.log");                                                // Create a log file based on the user logged into ADUtil
            Directory.CreateDirectory(BaseLogDirectory);                                                                         // Ensure the directory exist
            InitilizeLogFile();                                                                                                  // Set append mode.
            }// end of try
            catch (Exception ex)
            {
                Console.WriteLine($"Error validating log file: {ex.Message}");
            }// end of catch
        }// end of Auditlog manager constructor
        private void InitilizeLogFile()
        {
            File.AppendAllText(logFilePath, $"---------------------------------------------------------------------------------------------------------------------\n" +
                                            $"\t\t\t\tAudit log started at {DateTime.Now}\n" +
                                            $"---------------------------------------------------------------------------------------------------------------------\n");
        }// end of Initilizelogfile

        /// <summary>
        /// A method that log action to a log file in Logs folder.
        /// </summary>
        /// <param name="message"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void Log(string message)
        {
            if (string.IsNullOrEmpty(logFilePath))
            {
                throw new InvalidOperationException("Log file path is not set.");
            }// end of if statement

            string logEntry = $"{DateTime.Now}: {message}\n";
            try
            {
                File.AppendAllText(logFilePath, logEntry);                                                                     // Write the log message to the file.
            }// end of try
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing to log file: {ex.Message}");
            }// end of catch
        }// end of log

        /// <summary>
        /// a method that 
        /// </summary>
        public void RedirectConsoleOutput()
        {
            FileStream fileStream = new FileStream(logFilePath, FileMode.Append, FileAccess.Write);
            StreamWriter fileWriter = new StreamWriter(fileStream) { AutoFlush = true };

            TextWriter consoleWriter = Console.Out;
            TextWriter dualWriter = new DualWriterManager(consoleWriter, fileWriter);
            Console.SetOut(dualWriter);
            Console.SetError(dualWriter); // Optional: Redirect error output as well
        }// end of RedirectConsoleOutput
    }// end of class
}// end of namespace
