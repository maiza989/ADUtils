using Microsoft.Extensions.Configuration;
using Pastel;
using System.Collections;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace ADUtils
{
    /// <summary>
    /// A single remoting session to the on-prem Exchange server.
    ///
    /// Opens a runspace and creates a Microsoft.Exchange PSSession over WinRM; cmdlets are then
    /// dispatched into that session with Invoke-Command (see RunCommand) rather than imported
    /// locally. Extracted from AccountCreationManager so mailbox creation and shared-mailbox
    /// permissions share one implementation instead of duplicating the session plumbing.
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
        /// Opens the Exchange session. Returns false (with the reason printed) if the session could
        /// not be established -- callers must not report success when this fails.
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
                }

                // Deliberately NOT Import-PSSession. That cmdlet generates a temporary proxy script
                // module on disk and imports it, which is subject to execution policy -- and a
                // hosted PowerShell SDK runspace reports Restricted regardless of the machine's
                // LocalMachine setting, so it fails with "running scripts is disabled on this
                // system". Commands are dispatched with Invoke-Command -Session instead (see
                // RunCommand): nothing is written to disk, so no policy applies, and it skips
                // generating proxies for every Exchange cmdlet.

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
        /// Runs an Exchange cmdlet inside the remote session.
        ///
        /// Dispatched via <c>Invoke-Command -Session</c> and splatting rather than by calling the
        /// cmdlet directly, because without Import-PSSession the Exchange cmdlets do not exist in
        /// the local runspace -- they only exist on the server. See the note in Connect().
        /// </summary>
        /// <returns>True only when the cmdlet wrote nothing to the error stream.</returns>
        public bool RunCommand(string command, string description, Dictionary<string, object> parameters = null)
        {
            return RunQuery(command, description, parameters) != null;
        }// end of RunCommand

        /// <summary>
        /// Runs an Exchange cmdlet and returns its output.
        /// </summary>
        /// <returns>
        /// The cmdlet's output, which may be empty for a command that returns nothing, or null if
        /// the cmdlet failed. Callers must distinguish "empty" from "null".
        /// </returns>
        public Collection<PSObject> RunQuery(string command, string description, Dictionary<string, object> parameters = null)
        {
            if (_runspace == null || _session == null)
            {
                AppLog.Error($"Cannot run '{command}' — no Exchange session is open.");
                return null;
            }

            try
            {
                // Hashtable rather than Dictionary: it is what PowerShell splatting (@p) expects,
                // and it round-trips over remoting.
                var splat = new Hashtable();
                if (parameters != null)
                {
                    foreach (var parameter in parameters)
                    {
                        // A null value means a switch parameter; splatting expresses that as $true.
                        splat[parameter.Key] = parameter.Value ?? true;
                    }// end of foreach
                }

                using (PowerShell ps = PowerShell.Create())
                {
                    ps.Runspace = _runspace;
                    ps.AddCommand("Invoke-Command");
                    ps.AddParameter("Session", _session);
                    ps.AddParameter("ScriptBlock", ScriptBlock.Create("param($cmd, $p) & $cmd @p"));
                    ps.AddParameter("ArgumentList", new object[] { command, splat });

                    Collection<PSObject> results = ps.Invoke();
                    return ReportErrors(ps, description) ? null : results;
                }
            }
            catch (Exception ex)
            {
                AppLog.Error($"Error {description}: {ex.Message}", ex, Color.IndianRed);
                return null;
            }
        }// end of RunQuery

        /// <summary>
        /// Resolves whatever the operator typed -- SMTP address, alias or display name -- to the
        /// mailbox's distinguished name.
        ///
        /// Needed because the two cmdlet families disagree on identity: Add/Remove-MailboxPermission
        /// accepts an SMTP address, but Add/Remove-ADPermission works on AD objects and rejects one
        /// ("... wasn't found"). Resolving once to a DN makes both accept the same value, and fails
        /// early with a clear message if the mailbox does not exist.
        /// </summary>
        /// <returns>True when resolved; the DN and display name are set on success.</returns>
        public bool TryResolveMailbox(string mailbox, out string distinguishedName, out string displayName)
        {
            distinguishedName = null;
            displayName = null;

            var results = RunQuery("Get-Mailbox", $"looking up mailbox '{mailbox}'",
                                   new Dictionary<string, object> { ["Identity"] = mailbox });
            if (results == null) return false;

            if (results.Count == 0)
            {
                AppLog.Warn($"No mailbox matches '{mailbox}'.");
                return false;
            }

            distinguishedName = results[0].Properties["DistinguishedName"]?.Value?.ToString();
            displayName = results[0].Properties["Name"]?.Value?.ToString() ?? mailbox;

            if (string.IsNullOrWhiteSpace(distinguishedName))
            {
                AppLog.Warn($"Mailbox '{mailbox}' has no distinguished name — cannot set permissions on it.");
                return false;
            }
            return true;
        }// end of TryResolveMailbox

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
