using Microsoft.Extensions.Configuration;
using Pastel;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace ADUtils
{
    /// <summary>
    /// A single implicit-remoting session to the on-prem Exchange server.
    ///
    /// Opens a runspace, creates a Microsoft.Exchange PSSession over WinRM and imports it, so
    /// Exchange cmdlets can be invoked by name. Extracted from AccountCreationManager so mailbox
    /// creation and shared-mailbox permissions share one implementation instead of duplicating
    /// the session plumbing.
    ///
    /// Authenticates as the interactive user (Authentication = Default), which is how this has
    /// always worked -- the operator's Windows session must hold Exchange rights.
    /// </summary>
    public sealed class ExchangeSessionManager : IDisposable
    {
        private readonly string _exchangeServer;
        private readonly string _domainController;

        private Runspace _runspace;
        private PSObject _session;
        private bool _disposed;

        public string ExchangeDatabase { get; }

        /// <summary>The DC to pin Exchange cmdlets to, or null when none is configured.</summary>
        public string DomainController => string.IsNullOrWhiteSpace(_domainController) ? null : _domainController;

        public ExchangeSessionManager(IConfiguration configuration)
        {
            _exchangeServer = configuration["AccountCreationSettings:myExchangeServer"];
            _domainController = configuration["AccountCreationSettings:myDomainController"];
            ExchangeDatabase = configuration["AccountCreationSettings:myExchangeDatabase"];
        }// end of constructor

        /// <summary>
        /// Opens and imports the Exchange session. Returns false (with the reason printed) if the
        /// session could not be established -- callers must not report success when this fails.
        /// </summary>
        public bool Connect()
        {
            if (string.IsNullOrWhiteSpace(_exchangeServer))
            {
                AppLog.Error("No Exchange server configured (AccountCreationSettings:myExchangeServer).");
                return false;
            }

            try
            {
                _runspace = RunspaceFactory.CreateRunspace();
                _runspace.Open();

                using (PowerShell ps = PowerShell.Create())
                {
                    ps.Runspace = _runspace;

                    // No Set-ExecutionPolicy here: commands are invoked via AddCommand rather than
                    // script text, and the old call attempted a machine-wide policy change.
                    ps.AddCommand("New-PSSession");
                    ps.AddParameter("ConfigurationName", "Microsoft.Exchange");
                    ps.AddParameter("ConnectionUri", new Uri($"http://{_exchangeServer}/PowerShell/"));
                    ps.AddParameter("Authentication", "Default");

                    Collection<PSObject> result = ps.Invoke();
                    if (ReportErrors(ps, "creating the Exchange PSSession")) return false;

                    if (result == null || result.Count == 0)
                    {
                        AppLog.Error($"New-PSSession to '{_exchangeServer}' returned no session.");
                        return false;
                    }
                    _session = result[0];

                    ps.Commands.Clear();
                    ps.AddCommand("Import-PSSession");
                    ps.AddParameter("Session", _session);
                    ps.AddParameter("DisableNameChecking");
                    ps.AddParameter("AllowClobber");
                    ps.Invoke();
                    if (ReportErrors(ps, "importing the Exchange PSSession")) return false;
                }

                // Exchange cmdlets fail with "no available global catalog" when the session is
                // scoped to a single site. Widening the view to the whole forest is the documented
                // fix; it is best-effort because a single-domain forest works without it.
                if (!RunCommand("Set-ADServerSettings", "widening Exchange scope to the whole forest",
                                new Dictionary<string, object> { ["ViewEntireForest"] = true }))
                {
                    AppLog.Warn("Continuing without forest-wide scope.");
                }

                return true;
            }
            catch (Exception ex)
            {
                AppLog.Error($"Could not connect to Exchange at '{_exchangeServer}': {ex.Message}", ex, Color.IndianRed);
                return false;
            }
        }// end of Connect

        /// <summary>
        /// Invokes an Exchange cmdlet in the imported session.
        /// </summary>
        /// <returns>True only when the cmdlet wrote nothing to the error stream.</returns>
        public bool RunCommand(string command, string description, Dictionary<string, object> parameters = null)
        {
            if (_runspace == null)
            {
                AppLog.Error($"Cannot run '{command}' — no Exchange session is open.");
                return false;
            }

            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.Runspace = _runspace;
                    ps.AddCommand(command);

                    if (parameters != null)
                    {
                        foreach (var parameter in parameters)
                        {
                            if (parameter.Value == null)
                            {
                                ps.AddParameter(parameter.Key);      // switch parameter
                            }
                            else
                            {
                                ps.AddParameter(parameter.Key, parameter.Value);
                            }
                        }// end of foreach
                    }

                    ps.Invoke();
                    return !ReportErrors(ps, description);
                }
            }
            catch (Exception ex)
            {
                AppLog.Error($"Error {description}: {ex.Message}", ex, Color.IndianRed);
                return false;
            }
        }// end of RunCommand

        /// <summary>
        /// Prints anything on the PowerShell error stream.
        /// </summary>
        /// <returns>True if there were errors.</returns>
        private static bool ReportErrors(PowerShell ps, string description)
        {
            if (ps.Streams.Error.Count == 0) return false;

            foreach (ErrorRecord error in ps.Streams.Error)
            {
                // Log the exception object where PowerShell gave us one, so the file keeps the
                // stack trace; the console only ever needs the message.
                AppLog.Error($"Error {description}: {error.Exception?.Message ?? error.ToString()}", error.Exception);
            }// end of foreach
            ps.Streams.Error.Clear();
            return true;
        }// end of ReportErrors

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                if (_runspace != null && _session != null)
                {
                    using (PowerShell ps = PowerShell.Create())
                    {
                        ps.Runspace = _runspace;
                        ps.AddCommand("Remove-PSSession").AddParameter("Session", _session);
                        ps.Invoke();
                    }
                }
            }
            catch
            {
                // Tearing down a session that is already gone is not worth reporting.
            }

            _runspace?.Dispose();      // Dispose alone; Close() first would be a double teardown.
            _runspace = null;
            _session = null;
        }// end of Dispose
    }// end of class
}// end of namespace
