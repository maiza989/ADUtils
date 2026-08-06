using Microsoft.Extensions.Configuration;
using NLog;
using Pastel;
using System.Drawing;

namespace ADUtils
{
    /// <summary>
    /// The audit trail for privileged Active Directory changes.
    ///
    /// Backed by NLog (see NLog.config) rather than hand-rolled File.AppendAllText. That change
    /// fixed three problems with the previous implementation: a log directory that was unreachable
    /// left auditing silently disabled, there was no file locking so two instances could interleave
    /// or truncate each other, and the file grew without bound.
    ///
    /// Log() keeps its original signature so callers are unaffected.
    /// </summary>
    public class AuditLogManager
    {
        // Must match the logger name the audit rule targets in NLog.config.
        private static readonly Logger AuditLogger = LogManager.GetLogger("ADUtils.Audit");
        private static readonly Logger Diagnostics = LogManager.GetCurrentClassLogger();

        private readonly string _adminUsername;

        /// <summary>
        /// Where log files are written. Reported at startup so the operator knows where to look.
        /// </summary>
        public static string LogDirectory => Path.Combine(AppContext.BaseDirectory, "logs");

        /// <summary>
        /// Confirms NLog actually loaded a usable configuration, and reports it if not.
        ///
        /// NLog is configured not to throw on a bad config -- a logging problem must not take down
        /// a tool that makes privileged directory changes -- so without this check a rejected
        /// NLog.config would leave auditing silently disabled, which is exactly the failure the
        /// move to NLog was meant to eliminate.
        /// </summary>
        /// <returns>True when logging is working.</returns>
        public static bool VerifyLoggingConfigured()
        {
            var config = LogManager.Configuration;
            bool ok = config != null && config.AllTargets.Count > 0 && config.LoggingRules.Count > 0;

            if (!ok)
            {
                Console.WriteLine("WARNING: logging is NOT configured — no audit trail will be written.".Pastel(Color.Crimson));
                Console.WriteLine($"Check that NLog.config sits next to the executable, then read " +
                                  $"{Path.Combine(LogDirectory, "nlog-internal.log").Pastel(Color.MediumPurple)} for the reason.".Pastel(Color.Gray));
            }
            return ok;
        }// end of VerifyLoggingConfigured

        /// <param name="adminUsername">Stamped on every audit line as the actor.</param>
        /// <param name="configuration">
        /// Reserved for future logging settings. Log location now comes from NLog.config
        /// (${basedir}/logs) rather than LoggingSettings:BaseLogDirectory, so the logs travel with
        /// the executable instead of depending on a mapped drive being present.
        /// </param>
        public AuditLogManager(string adminUsername, IConfiguration configuration)
        {
            _adminUsername = string.IsNullOrWhiteSpace(adminUsername) ? "unknown" : adminUsername;

            AuditLogger.WithProperty("admin", _adminUsername).Info("=== ADUtils session started ===");
            Diagnostics.Info("Session started for {0} on {1} (logs: {2})", _adminUsername, Environment.MachineName, LogDirectory);
        }// end of AuditLogManager constructor

        /// <summary>
        /// Records a completed privileged action. Call this only after the change actually
        /// succeeded -- these lines are the record that it happened.
        /// </summary>
        public void Log(string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return;

            // Collapse the multi-line entries the managers build so one action is one audit line.
            string singleLine = message.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' ').Trim();
            while (singleLine.Contains("  ")) singleLine = singleLine.Replace("  ", " ");

            // NLog never throws out of a write, so a logging failure can no longer take down or
            // silently disable an audited operation.
            AuditLogger.WithProperty("admin", _adminUsername).Info(singleLine);
        }// end of Log

        /// <summary>
        /// Records the session ending and flushes buffered writes. Targets are async, so without
        /// the flush the last entries can be lost when the process exits.
        /// </summary>
        public void EndSession()
        {
            AuditLogger.WithProperty("admin", _adminUsername).Info("=== ADUtils session ended ===");
            LogManager.Flush();
        }// end of EndSession
    }// end of class
}// end of namespace
