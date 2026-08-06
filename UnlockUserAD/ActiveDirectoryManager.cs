using Microsoft.Extensions.Configuration;
using Pastel;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;

namespace ADUtils
{

    public class ActiveDirectoryManager
    {
        /// <summary>
        /// Budget for the whole WinRM lockout query across all reachable DCs.
        ///
        /// Generous on purpose: a Security log scan on a busy DC can take a while, and the cost of
        /// setting this too low is a spurious "gave up" on a query that would have succeeded.
        /// Unreachable DCs are already excluded by a fast port probe before this applies, and
        /// genuine transport or authorization failures return in well under a second.
        /// </summary>
        private const int LockoutLookupTimeoutSeconds = 30;

        /// <summary>
        /// How far back to search for the lockout event. Bounding this keeps the DC from scanning
        /// the entire Security log.
        /// </summary>
        private const int LockoutSearchWindowDays = 7;

        PasswordManager passwordManager = new PasswordManager();

        private string _adminUsername;
        private string _adminPassword;
        private string _domain;
        private List<string> _domainControllers = new List<string>();

        /// <summary>
        /// Set once the remote Security log is shown to be genuinely unreachable or the credentials
        /// rejected, so later lookups in the same session fail instantly instead of re-probing.
        ///
        /// Deliberately not set by a query timeout — that is ambiguous, and caching it would
        /// disable the feature for the session over one slow query.
        /// </summary>
        private bool _eventLogBlocked;

        public void SetAdminCredentials(string adminUsername, string adminPassword, IConfiguration configuration)
        {
            _adminUsername = adminUsername;
            _adminPassword = adminPassword;

            // myDomainName holds a *domain controller* hostname (e.g. "LMDC2"), which is what
            // PrincipalContext wants. The WinRM credential needs the account's DOMAIN to qualify
            // the username, so read myDomain instead.
            _domain = configuration["AccountCreationSettings:myDomain"];

            // Load all domain controllers from config — uses GetChildren() to avoid requiring Binder package
            _domainControllers = configuration.GetSection("AccountCreationSettings:myDomainControllers")
                                              .GetChildren()
                                              .Select(c => c.Value)
                                              .Where(v => !string.IsNullOrWhiteSpace(v))
                                              .ToList();

            if (_domainControllers.Count == 0)
                AppLog.Warn("Warning: No domain controllers configured in Appsettings.json.", color: Color.DarkGoldenrod);
        }// end of SetAdminCredentials

        /// <summary>
        /// A method that display a general information about a user.
        /// </summary>
        /// <param name="context">The PrincipalContext to use for querying Active Directory</param>
        public void DisplayUserInfo(PrincipalContext context)
        {
            bool returnToMenu = false;
            List<string> userGroups = new List<string>();
            do
            {
                AppLog.Prompt($"Enter the username to display info about (Type {"'exit'".Pastel(Color.MediumPurple)} to return to the main menu): ");
                string username = ConsoleInput.ReadTrimmedLower();

                if (username == "exit")
                {
                    returnToMenu = true;
                }// end of if statement
                else if (username.Length == 0)
                {
                    AppLog.Warn("Enter a username, or 'exit' to return to the main menu.", color: Color.DarkGoldenrod);
                }
                else
                {
                    // This method had no exception handling at all, so any AD hiccup escaped to
                    // Main's catch, which resets isAuthenticated and re-prompts for credentials --
                    // a read-only lookup should never log the operator out.
                    try
                    {
                        UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
                        if (user != null)
                        {
                            var groups = user.GetGroups();
                            foreach (var group in groups)
                            {
                                userGroups.Add(group.Name);
                            }// end of foreach
                            userGroups.Sort();
                            string userGroupsString = userGroups.Count > 0 ? string.Join(", ", userGroups) : "N/A";

                            string title = "N/A";
                            string department = "N/A";
                            if (user.GetUnderlyingObject() is DirectoryEntry directoryEntry)
                            {
                                title = directoryEntry.Properties["title"].Value as string ?? "N/A";
                                department = directoryEntry.Properties["department"].Value as string ?? "N/A";
                            }

                            string lastBadPwd = user.LastBadPasswordAttempt.HasValue
                                ? TimeZoneInfo.ConvertTimeFromUtc(user.LastBadPasswordAttempt.Value.ToUniversalTime(), TimeZoneInfo.Local).ToString()
                                : "N/A";

                            // Note: LastLogon reads the non-replicated per-DC lastLogon attribute, so
                            // the answer varies by which DC responded.
                            string lastLogon = user.LastLogon.HasValue
                                ? TimeZoneInfo.ConvertTimeFromUtc(user.LastLogon.Value.ToUniversalTime(), TimeZoneInfo.Local).ToString()
                                : "N/A";

                            DateTime expires = passwordManager.GetPasswordExpirationDate(user);
                            bool locked = user.IsAccountLockedOut();

                            ConsoleUi.Panel($"{user.SamAccountName} {(user.DisplayName != null ? "· " + user.DisplayName : "")}".Trim(), new[]
                            {
                                ("First name", user.GivenName ?? "N/A"),
                                ("Last name", user.Surname ?? "N/A"),
                                ("Email", user.EmailAddress ?? "N/A"),
                                ("Title", title),
                                ("Department", department),
                                ("Account", ConsoleUi.State(user.Enabled == true, user.Enabled == true ? "Enabled" : "DISABLED")),
                                ("Lockout", ConsoleUi.State(!locked, locked ? "LOCKED" : "Not locked")),
                                ("Password set", passwordManager.GetPasswordLastSetDate(user)?.ToString("yyyy-MM-dd HH:mm") ?? "N/A"),
                                ("Password expires", user.PasswordNeverExpires ? "never"
                                                     : expires == DateTime.MinValue ? "unknown"
                                                     : expires.ToString("yyyy-MM-dd")),
                                ("Bad logon count", user.BadLogonCount.ToString()),
                                ("Last logon", lastLogon),
                                ("Last bad logon", lastBadPwd),
                                ("Home directory", user.HomeDirectory ?? "N/A"),
                                ("SID", user.Sid?.ToString() ?? "N/A"),
                                ("Member of", userGroupsString)
                            });
                        }// end of if statement
                        else
                        {
                            // Previously this branch was missing entirely, so a typo printed nothing.
                            AppLog.Warn($"\tUser '{username}' not found in Active Directory.", color: Color.IndianRed);
                        }// end of else statement
                    }// end of try
                    catch (Exception ex)
                    {
                        AppLog.Error($"Error looking up '{username}': {ex.Message}", ex, Color.IndianRed);
                    }// end of catch
                }// end of else statement
                userGroups.Clear();
            } while (!returnToMenu);
        }// end of DisplayUserInfo
        /// <summary>
        /// A method to unlock one sepcific user.
        /// </summary>
        /// <param name="context">Based in what the computer domain</param>
        public void UnlockUser(PrincipalContext context)
        {
            bool returnToMenu = false;
            do
            {
                AppLog.Prompt($"Enter the username to unlock (type {"'exit'".Pastel(Color.MediumPurple)} to return to the main menu): ");
                string username = ConsoleInput.ReadTrimmedLower();

                if (username.ToLower().Trim() == "exit")
                {
                    returnToMenu = true;
                }// end of if statement
                else
                {
                    try
                    {
                        UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);                           // Search for specific user using username
                        if (user != null)
                        {
                            if (user.IsAccountLockedOut())                                                                                           // Check if the user is locked
                            {
                                user.UnlockAccount();                                                                                                // Unlock user
                                AppLog.Info($"\tUser account '{username}' has been unlocked.", Color.LimeGreen);
                            }// end of inner-if-statement
                            else
                            {
                                AppLog.Screen($"\tUser account '{username}' is not locked.", Color.OrangeRed);
                            }// end of else-statement
                        }// end of Outter-if-statement
                        else
                        {
                            AppLog.Warn($"\tUser account '{username}' not found in Active Directory.", color: Color.IndianRed);
                        }// end of else-statement
                    }// end of Try-Catch
                    catch (Exception ex)
                    {
                        AppLog.Error($"Error Unlocking a user: {ex.Message}", ex, Color.IndianRed);
                    }// end of catch
                }// end of else
            } while (!returnToMenu);
        }// end of UnlockUser

        /// <summary>
        /// A method to go through every user in Active Directory and unlock all of them if any is locked.
        /// </summary>
        /// <param name="context">Based in what the computer domain</param>
        public void UnlockAllUsers(PrincipalContext context)
        {
            try
            {
                AppLog.Screen("\nUnlocking all user accounts...");
                bool anyUnlocked = false;
                using (PrincipalSearcher searcher = new PrincipalSearcher(new UserPrincipal(context) { Enabled = true }))
                using (var results = searcher.FindAll())
                {
                    foreach (var result in results)
                    {
                        using (result)
                        {
                            // The second identical test that used to follow this was always true.
                            if (!(result is UserPrincipal user) || !user.IsAccountLockedOut()) continue;

                            try
                            {
                                user.UnlockAccount();
                                AppLog.Info($"\t[{DateTime.Now:MM-dd-yyyy HH:mm:ss tt}]: User account '{user.SamAccountName}' has been unlocked.", Color.LimeGreen);
                                anyUnlocked = true;
                            }
                            catch (Exception ex)
                            {
                                // Report and keep going -- one account we lack rights on shouldn't
                                // abandon the rest of the sweep.
                                AppLog.Error($"\tCould not unlock '{user.SamAccountName}': {ex.Message}", ex, Color.IndianRed);
                            }
                        }
                    }// end of foreach
                }
                if (!anyUnlocked)                                                                                                              // If-Else statement to check if any user were unlocked and print appropriate response.
                {
                    AppLog.Warn("\tNo user accounts were locked.", color: Color.DarkGoldenrod);
                }// end of if-statement
                else
                {
                    AppLog.Info("\nAll user accounts have been unlocked successfully.", Color.DarkCyan);
                }// end of else-statement
            }// end of Try-Catch
            catch (Exception ex)
            {
                AppLog.Error($"Error: {ex.Message}", ex, Color.IndianRed);
            }// end of Catch
        }// end of UnlockAllUsers

        // TODO - grab the lockout event
        /*
                private void PrintLockoutEventDetails(string username)
                {
                    string query = $"*[System/EventID=4740] and *[EventData[Data[@Name='TargetUserName'] and (Data='{username}')]]";
                    EventLogQuery eventsQuery = new EventLogQuery("Security", PathType.LogName, query);

                    try
                    {
                        using (EventLogReader logReader = new EventLogReader(eventsQuery))
                        {
                            EventRecord eventRecord = logReader.ReadEvent();
                            if (eventRecord != null)
                            {
                                // Extract relevant details
                                string workstation = eventRecord.Properties[1].Value.ToString();
                                string lockedUser = eventRecord.Properties[0].Value.ToString();
                                DateTime? lockoutTime = eventRecord.TimeCreated;

                                Console.WriteLine($"\tAccount: {lockedUser} was locked on workstation: {workstation} at {lockoutTime}".Pastel(Color.Gold));
                            }
                            else
                            {
                                Console.WriteLine($"\tNo lockout events found for {username}.".Pastel(Color.OrangeRed));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error fetching lockout details: {ex.Message}".Pastel(Color.IndianRed));
                    }
                }
        */

   /*     /// <summary>
        /// Reads Security Event Log for Event ID 4740 and returns the CallerComputerName
        /// </summary>
        private string GetWorkstationNameFromEvent(string username)
        {
            string domainController = "DC01.domain.local"; // ✅ Replace with your actual DC
            if (string.IsNullOrEmpty(_adminUsername) || string.IsNullOrEmpty(_adminPassword))
            {
                Console.WriteLine("Admin credentials are missing.");
                return "Unknown";
            }

            try
            {
                SecureString securePassword = new SecureString();
                foreach (char c in _adminPassword)
                    securePassword.AppendChar(c);
                securePassword.MakeReadOnly();

                using (var session = new EventLogSession(
                           domainController,
                           _domain,
                           _adminUsername,
                           securePassword,
                           SessionAuthentication.Default))
                {
                    string query = "*[System[EventID=4740]]";
                    EventLogQuery eventsQuery = new EventLogQuery("Security", PathType.LogName, query)
                    {
                        Session = session
                    };

                    using (EventLogReader logReader = new EventLogReader(eventsQuery))
                    {
                        for (EventRecord evt = logReader.ReadEvent(); evt != null; evt = logReader.ReadEvent())
                        {
                            string xml = evt.ToXml();
                            if (xml.Contains(username, StringComparison.OrdinalIgnoreCase))
                            {
                                string callerName = ExtractCallerComputerName(xml);
                                evt.Dispose();
                                if (!string.IsNullOrEmpty(callerName))
                                    return callerName;
                            }
                            evt.Dispose();
                        }
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                Console.WriteLine($"Access denied while reading Security log on {domainController}. " +
                                  $"Ensure {_adminUsername} has 'Event Log Readers' permission or run as Domain Admin."
                                  .Pastel(Color.DarkOrange));
            }
            catch (EventLogNotFoundException ex)
            {
                Console.WriteLine($"Security log not found on {domainController}: {ex.Message}".Pastel(Color.IndianRed));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting workstation for {username}: {ex.Message}".Pastel(Color.DarkOrange));
            }

            return "Unknown";
        }


        /// <summary>
        /// Extracts the CallerComputerName value from event XML
        /// </summary>
        private string ExtractCallerComputerName(string xml)
        {
            try
            {
                var startTag = "<Data Name='CallerComputerName'>";
                var endTag = "</Data>";
                int startIndex = xml.IndexOf(startTag);
                if (startIndex == -1) return null;
                startIndex += startTag.Length;
                int endIndex = xml.IndexOf(endTag, startIndex);
                if (endIndex == -1) return null;
                return xml.Substring(startIndex, endIndex - startIndex).Trim();
            }
            catch
            {
                return null;
            }
        }*/


        /// <summary>
        /// A method to check if any user is locked in Active Directory.
        /// </summary>
        /// <param name="context">Based in what the computer domain</param>
        public void CheckLockedAccounts(PrincipalContext context)
        {
            AppLog.Screen("\nLocked user accounts:");
            try
            {
                // Collect the locked accounts first, so the event logs are only queried when
                // there is actually something to look up.
                var lockedUsers = new List<(string SamAccountName, DateTime? LockoutTime)>();

                using (PrincipalSearcher searcher = new PrincipalSearcher(new UserPrincipal(context) { Enabled = true }))                       // Creating the search object
                using (var results = searcher.FindAll())
                {
                    foreach (var result in results)                                                                                             // Look through what is in the user search object
                    {
                        using (result)
                        {
                            // The second identical test that used to follow this was always true.
                            if (!(result is UserPrincipal user) || !user.IsAccountLockedOut()) continue;

                            lockedUsers.Add((user.SamAccountName, ReadLockoutTime(user)));
                        }
                    }// end of foreach
                }

                if (lockedUsers.Count == 0)
                {
                    AppLog.Screen($"\tNo accounts are LOCKED!!! YAY!!!.", Color.RoyalBlue);
                    return;
                }

                // Deliberately no event-log lookup here. Listing who is locked must stay fast;
                // finding *where* they were locked costs seconds per DC and is a drill-down, so it
                // lives on its own menu option (see FindLockoutSource).
                foreach (var locked in lockedUsers)
                {
                    string when = locked.LockoutTime?.ToString("MM-dd-yyyy HH:mm:ss tt") ?? "time unavailable";

                    // Printed for every locked account. This line used to sit inside the
                    // "lockoutTime exists" branch, so accounts without a readable lockoutTime were
                    // counted as locked but never displayed.
                    AppLog.Warn($"\t[{when}] - {locked.SamAccountName}", color: Color.Crimson);
                }// end of foreach

                AppLog.Screen($"\n\t{lockedUsers.Count} locked account(s). Use {"'Find Lockout Source'".Pastel(Color.MediumPurple)} to see which machine locked one.", Color.Gray);
            }// end of try-catch
            catch (Exception ex)
            {
                AppLog.Error($"Error: {ex.Message}", ex, Color.IndianRed);
            }// end of catch
        }// end of CheckLockedAccounts

        /// <summary>
        /// Reads the COM lockoutTime attribute off a user, or null when it is absent/unreadable.
        /// </summary>
        private DateTime? ReadLockoutTime(UserPrincipal user)
        {
            try
            {
                if (!(user.GetUnderlyingObject() is DirectoryEntry directoryEntry)) return null;
                if (!directoryEntry.Properties.Contains("lockoutTime")) return null;

                object lockOutValue = directoryEntry.Properties["lockoutTime"].Value;
                if (lockOutValue == null) return null;

                // lockoutTime comes back as a COM IADsLargeInteger; read HighPart/LowPart by reflection.
                var highPart = (int)lockOutValue.GetType().InvokeMember("HighPart", System.Reflection.BindingFlags.GetProperty, null, lockOutValue, null);
                var lowPart = (int)lockOutValue.GetType().InvokeMember("LowPart", System.Reflection.BindingFlags.GetProperty, null, lockOutValue, null);
                long lockoutTicks = ((long)highPart << 32) + (uint)lowPart;

                return lockoutTicks > 0 ? DateTime.FromFileTimeUtc(lockoutTicks).ToLocalTime() : (DateTime?)null;
            }
            catch (Exception ex)
            {
                AppLog.Warn($"Error reading lockoutTime for {user.SamAccountName}: {ex.Message}", ex, Color.DarkGoldenrod);
                return null;
            }
        }// end of ReadLockoutTime

        /// <summary>
        /// Interactive drill-down: asks for a username and reports which machine locked it out,
        /// from Security Event ID 4740.
        ///
        /// Kept off the "Check All Locked Accounts" path on purpose -- reading a remote Security
        /// log costs seconds per domain controller, and doing it for every locked account made a
        /// simple listing take ~30 seconds.
        /// </summary>
        public void FindLockoutSource()
        {
            bool returnToMenu = false;
            do
            {
                AppLog.Prompt($"Enter the username to trace the lockout for (Type {"'exit'".Pastel(Color.MediumPurple)} to return to the menu): ");
                string username = ConsoleInput.ReadTrimmed();

                if (username.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    returnToMenu = true;
                }
                else if (username.Length == 0)
                {
                    AppLog.Warn("Enter a username, or 'exit' to return to the menu.", color: Color.DarkGoldenrod);
                }
                else if (username.IndexOfAny(new[] { '\'', '"', '[', ']', '(', ')', '\\', '/' }) >= 0)
                {
                    // The name is interpolated into an XPath predicate below.
                    AppLog.Warn("That is not a valid sAMAccountName.", color: Color.IndianRed);
                }
                else
                {
                    string source = LookupLockoutSource(username);
                    if (source != null)
                    {
                        AppLog.Info($"\t'{username}' was locked out from: {source.Pastel(Color.Gold)}", Color.LimeGreen);
                    }
                    else
                    {
                        AppLog.Warn($"\tNo lockout source found for '{username}'.");
                    }
                }
            } while (!returnToMenu);
        }// end of FindLockoutSource

        /// <summary>
        /// Reads Security Event ID 4740 to find the workstation that locked out one account.
        ///
        /// Queries the PDC emulator first, since it processes lockouts and therefore reliably holds
        /// the 4740 record; other configured DCs are only tried as a fallback. Filtering happens
        /// server-side via XPath on TargetUserName, so each DC returns one event instead of this
        /// pulling hundreds and scanning them here. Unreachable DCs are skipped by a short TCP
        /// probe rather than by waiting out the RPC timeout.
        /// </summary>
        /// <returns>The workstation name, or null when it could not be determined.</returns>
        private string LookupLockoutSource(string username)
        {
            if (_eventLogBlocked)
            {
                // Established earlier this session -- don't spend the timeout again.
                PrintEventLogBlockedHelp();
                return null;
            }

            if (string.IsNullOrEmpty(_adminUsername) || string.IsNullOrEmpty(_adminPassword))
            {
                AppLog.Warn("Admin credentials not set — cannot read the Security log.", color: Color.DarkGoldenrod);
                return null;
            }

            var candidates = BuildDomainControllerSearchOrder();
            if (candidates.Count == 0)
            {
                AppLog.Warn("No domain controllers could be discovered or configured.", color: Color.DarkGoldenrod);
                return null;
            }

            // Drop hosts that aren't listening for WinRM before asking PowerShell to connect --
            // an unreachable DC costs ~700ms here instead of the remoting stack's own timeout.
            // Probed concurrently, so the cost is one timeout rather than one per dead DC; the
            // results are still reported in PDC-first order for predictable output.
            var probes = candidates.Select(dc => new { Dc = dc, Open = Task.Run(() => IsWinRmPortOpen(dc)) }).ToList();
            Task.WaitAll(probes.Select(p => p.Open).ToArray());

            var reachable = new List<string>();
            foreach (var probe in probes)
            {
                if (probe.Open.Result) reachable.Add(probe.Dc);
                else AppLog.Screen($"\t{probe.Dc}: WinRM not reachable — skipped.", Color.DarkGray);
            }

            if (reachable.Count == 0)
            {
                AppLog.Warn("\tNone of the domain controllers are reachable over WinRM.", color: Color.IndianRed);
                _eventLogBlocked = true;
                PrintEventLogBlockedHelp();
                return null;
            }

            // Filter server-side on both the event id and the target account, and cap the search to
            // a recent window. Without the time bound the DC scans the whole Security log, which on
            // a busy DC is far slower than the lookup is worth.
            long windowMs = (long)TimeSpan.FromDays(LockoutSearchWindowDays).TotalMilliseconds;
            string xpath = $"*[System[EventID=4740 and TimeCreated[timediff(@SystemTime) <= {windowMs}]]]" +
                           $" and *[EventData[Data[@Name='TargetUserName']='{username}']]";

            AppLog.Screen($"\tQuerying {reachable.Count} DC(s) over WinRM: {string.Join(", ", reachable)}", Color.DarkCyan);

            // Bound the wait rather than letting the remoting stack decide how long the operator
            // sits there. The abandoned task unwinds on its own; nothing downstream depends on it.
            var lookup = Task.Run(() => QueryLockoutEvent(reachable, xpath));

            if (!lookup.Wait(TimeSpan.FromSeconds(LockoutLookupTimeoutSeconds)))
            {
                // Deliberately does NOT set _eventLogBlocked. A timeout is ambiguous -- it can mean
                // a slow Security log scan on a busy DC just as easily as a hung transport -- and
                // caching it would disable the feature for the rest of the session over what may
                // have been one slow query. Real transport and authorization failures come back in
                // well under a second, and those do get cached.
                AppLog.Warn($"\tGave up after {LockoutLookupTimeoutSeconds}s — the DC did not answer in time. Try again.", color: Color.DarkGoldenrod);
                return null;
            }

            LockoutQueryResult result = lookup.Result;

            if (result.Hit != null)
            {
                AppLog.Info($"\t{result.Hit.DomainController} answered: locked out {result.Hit.TimeCreated:MM-dd-yyyy HH:mm:ss}.", Color.DarkOliveGreen);
                return result.Hit.Caller;
            }

            if (result.Failure != null)
            {
                // A transport or authorization failure applies to every DC equally, so remember it
                // and stop retrying for the rest of the session. A clean "no event for this user"
                // leaves _eventLogBlocked alone -- that is a normal negative result, and poisoning
                // the cache with it would break every later lookup.
                AppLog.Warn($"\t{result.Failure}", color: Color.IndianRed);
                _eventLogBlocked = true;
                PrintEventLogBlockedHelp();
                return null;
            }

            AppLog.Screen($"\tNo Event 4740 for '{username}' on any DC in the last {LockoutSearchWindowDays} days.", Color.DarkGray);
            return null;
        }// end of LookupLockoutSource

        /// <summary>One DC's answer for a lockout query.</summary>
        private sealed class LockoutHit
        {
            public string Caller { get; init; }
            public DateTime TimeCreated { get; init; }
            public string DomainController { get; init; }
        }

        /// <summary>
        /// Outcome of a lockout query. Distinguishes the three cases the caller must treat
        /// differently: found, genuinely nothing recorded, and could-not-ask.
        /// </summary>
        private sealed class LockoutQueryResult
        {
            public LockoutHit Hit { get; init; }

            /// <summary>Set only when the query could not be performed, not when it found nothing.</summary>
            public string Failure { get; init; }
        }

        /// <summary>
        /// Runs the 4740 query against every DC in one Invoke-Command.
        ///
        /// WinRM rather than <c>EventLogSession</c>: the event-log RPC interface is blocked in this
        /// environment (port 135 answers but the dynamic range it redirects to does not), and that
        /// is pre-authentication so no credential helps. WinRM is reachable and the admin account
        /// is a Domain Admin, so it is authorized to read a DC's Security log.
        ///
        /// PowerShell fans -ComputerName out in parallel itself, so this is one round trip for all
        /// DCs rather than a sequential loop.
        /// </summary>
        private LockoutQueryResult QueryLockoutEvent(List<string> domainControllers, string xpath)
        {
            try
            {
                using Runspace runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();

                using PowerShell ps = PowerShell.Create();
                ps.Runspace = runspace;

                // -ErrorAction SilentlyContinue is required, not cosmetic: Get-WinEvent raises a
                // *terminating* error when nothing matches the filter, which must read as
                // "not found" rather than as a failure.
                // TimeCreated plus the raw XML lets ExtractXmlDataValue below do the parsing,
                // keeping the remote script minimal.
                ps.AddCommand("Invoke-Command");
                ps.AddParameter("ComputerName", domainControllers.ToArray());
                ps.AddParameter("Credential", BuildAdminCredential());
                ps.AddParameter("ScriptBlock", ScriptBlock.Create(@"
                    Get-WinEvent -LogName Security -FilterXPath $args[0] -MaxEvents 1 -ErrorAction SilentlyContinue |
                        Select-Object TimeCreated, @{ n = 'Xml'; e = { $_.ToXml() } }"));
                ps.AddParameter("ArgumentList", new object[] { xpath });

                var results = ps.Invoke();

                if (ps.Streams.Error.Count > 0)
                {
                    // Report the first error but keep any results: one unreachable DC in the set
                    // shouldn't discard an answer another DC returned.
                    string first = ps.Streams.Error[0].Exception?.Message ?? ps.Streams.Error[0].ToString();
                    foreach (var e in ps.Streams.Error)
                    {
                        AppLog.Detail($"WinRM error: {e.Exception?.Message ?? e.ToString()}");
                    }
                    if (results == null || results.Count == 0)
                    {
                        return new LockoutQueryResult { Failure = first };
                    }
                }

                if (results == null || results.Count == 0) return new LockoutQueryResult();

                // More than one DC can hold an event for the same account; the newest wins.
                LockoutHit best = null;
                foreach (PSObject result in results)
                {
                    string xml = result.Properties["Xml"]?.Value as string;
                    if (string.IsNullOrWhiteSpace(xml)) continue;

                    string caller = ExtractXmlDataValue(xml, "CallerComputerName");
                    if (string.IsNullOrWhiteSpace(caller)) continue;

                    DateTime when = result.Properties["TimeCreated"]?.Value is DateTime dt ? dt : DateTime.MinValue;
                    string dc = result.Properties["PSComputerName"]?.Value as string ?? "unknown DC";

                    if (best == null || when > best.TimeCreated)
                    {
                        best = new LockoutHit
                        {
                            Caller = caller.TrimStart('\\'),
                            TimeCreated = when,
                            DomainController = dc
                        };
                    }
                }
                return new LockoutQueryResult { Hit = best };
            }
            catch (Exception ex)
            {
                return new LockoutQueryResult { Failure = ex.Message };
            }
        }// end of QueryLockoutEvent

        /// <summary>One Event 4740 record.</summary>
        private sealed class LockoutEvent
        {
            public DateTime When { get; init; }
            public string Account { get; init; }
            public string Caller { get; init; }
            public string DomainController { get; init; }
        }

        /// <summary>
        /// All Event 4740 records from the last N hours across every reachable DC.
        /// Same transport as the single-user lookup, without the TargetUserName filter.
        /// </summary>
        /// <returns>The events, or null when the log could not be read at all.</returns>
        private List<LockoutEvent> QueryRecentLockoutEvents(int hours)
        {
            if (!AdminSession.IsSet)
            {
                ConsoleUi.Warn("Admin credentials not set — cannot read the Security log.");
                return null;
            }

            var reachable = GetReachableDomainControllers();
            if (reachable.Count == 0) return null;

            long windowMs = (long)TimeSpan.FromHours(hours).TotalMilliseconds;
            string xpath = $"*[System[EventID=4740 and TimeCreated[timediff(@SystemTime) <= {windowMs}]]]";

            ConsoleUi.Note($"Querying {reachable.Count} DC(s) for lockouts in the last {hours}h: {string.Join(", ", reachable)}");

            try
            {
                using Runspace runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();
                using PowerShell ps = PowerShell.Create();
                ps.Runspace = runspace;

                ps.AddCommand("Invoke-Command");
                ps.AddParameter("ComputerName", reachable.ToArray());
                ps.AddParameter("Credential", AdminSession.CreatePsCredential());
                ps.AddParameter("ScriptBlock", ScriptBlock.Create(@"
                    Get-WinEvent -LogName Security -FilterXPath $args[0] -MaxEvents 500 -ErrorAction SilentlyContinue |
                        Select-Object TimeCreated, @{ n = 'Xml'; e = { $_.ToXml() } }"));
                ps.AddParameter("ArgumentList", new object[] { xpath });

                var results = ps.Invoke();

                foreach (var e in ps.Streams.Error)
                {
                    AppLog.Detail($"WinRM error: {e.Exception?.Message ?? e.ToString()}");
                }

                if (results == null || results.Count == 0)
                {
                    if (ps.Streams.Error.Count > 0)
                    {
                        _eventLogBlocked = true;
                        PrintEventLogBlockedHelp();
                        return null;
                    }
                    return new List<LockoutEvent>();
                }

                var events = new List<LockoutEvent>();
                foreach (PSObject result in results)
                {
                    string xml = result.Properties["Xml"]?.Value as string;
                    if (string.IsNullOrWhiteSpace(xml)) continue;

                    string account = ExtractXmlDataValue(xml, "TargetUserName");
                    string caller = ExtractXmlDataValue(xml, "CallerComputerName");
                    if (string.IsNullOrWhiteSpace(account)) continue;

                    events.Add(new LockoutEvent
                    {
                        When = result.Properties["TimeCreated"]?.Value is DateTime dt ? dt : DateTime.MinValue,
                        Account = account,
                        Caller = string.IsNullOrWhiteSpace(caller) ? "unknown" : caller.TrimStart('\\'),
                        DomainController = result.Properties["PSComputerName"]?.Value as string ?? "?"
                    });
                }
                return events;
            }
            catch (Exception ex)
            {
                ConsoleUi.Fail($"Could not read lockout events: {ex.Message}", ex);
                return null;
            }
        }// end of QueryRecentLockoutEvents

        /// <summary>
        /// Configured/discovered DCs filtered to those answering on the WinRM port, probed
        /// concurrently so one dead DC costs a single timeout rather than one each.
        /// </summary>
        private List<string> GetReachableDomainControllers()
        {
            var candidates = BuildDomainControllerSearchOrder();
            if (candidates.Count == 0)
            {
                ConsoleUi.Warn("No domain controllers could be discovered or configured.");
                return new List<string>();
            }

            var probes = candidates.Select(dc => new { Dc = dc, Open = Task.Run(() => IsWinRmPortOpen(dc)) }).ToList();
            Task.WaitAll(probes.Select(p => p.Open).ToArray());

            var reachable = new List<string>();
            foreach (var probe in probes)
            {
                if (probe.Open.Result) reachable.Add(probe.Dc);
                else ConsoleUi.Note($"{probe.Dc}: WinRM not reachable — skipped.");
            }

            if (reachable.Count == 0)
            {
                ConsoleUi.Fail("None of the domain controllers are reachable over WinRM.");
                _eventLogBlocked = true;
                PrintEventLogBlockedHelp();
            }
            return reachable;
        }// end of GetReachableDomainControllers

        /// <summary>
        /// Builds the PSCredential for remoting from the credentials captured at login.
        ///
        /// The name has to be domain-qualified for WinRM, but must not be qualified twice if the
        /// operator already typed DOMAIN\user or a UPN. Note the password is not sent in the clear
        /// on port 5985 -- PowerShell remoting encrypts the payload via Kerberos/Negotiate.
        /// </summary>
        private PSCredential BuildAdminCredential()
        {
            string user = _adminUsername.Contains('\\') || _adminUsername.Contains('@') || string.IsNullOrWhiteSpace(_domain)
                ? _adminUsername
                : $"{_domain}\\{_adminUsername}";

            SecureString securePassword = new SecureString();
            foreach (char c in _adminPassword)
                securePassword.AppendChar(c);
            securePassword.MakeReadOnly();

            return new PSCredential(user, securePassword);
        }// end of BuildAdminCredential

        /// <summary>
        /// Explains what has to change for this lookup to work.
        /// </summary>
        private static void PrintEventLogBlockedHelp()
        {
            AppLog.Warn("\tCould not read the Security log from any domain controller.", color: Color.IndianRed);
            AppLog.Screen($"\tThis lookup uses {"WinRM".Pastel(Color.MediumPurple)} (TCP 5985). Check, on the DCs:", Color.Gray);
            AppLog.Screen($"\t  - WinRM is enabled and its firewall rule allows this subnet  ({"Test-WSMan -ComputerName <dc>".Pastel(Color.MediumPurple)})", Color.Gray);
            AppLog.Screen("\t  - the admin account you logged in with is an administrator on the DC", Color.Gray);
            AppLog.Screen("\tUntil then, trace lockouts from Event Viewer on the DC (Security log, Event ID 4740).", Color.Gray);
        }// end of PrintEventLogBlockedHelp

        /// <summary>
        /// The DCs to query, PDC emulator first because it processes lockouts and so is the most
        /// likely to hold the 4740 record.
        ///
        /// Discovered from AD rather than hardcoded: the configured list had gone stale in both
        /// directions -- it named a decommissioned DC and omitted a live one, so lockouts recorded
        /// by the missing DC could never be found. <c>myDomainControllers</c> is now an optional
        /// restriction for when the search should be limited to specific DCs.
        /// </summary>
        private List<string> BuildDomainControllerSearchOrder()
        {
            // Explicit restriction wins when configured.
            if (_domainControllers.Count > 0)
            {
                AppLog.Detail($"Using the configured domain controller list: {string.Join(", ", _domainControllers)}");
                return OrderPdcFirst(_domainControllers);
            }

            try
            {
                var discovered = System.DirectoryServices.ActiveDirectory.Domain.GetComputerDomain()
                                     .DomainControllers
                                     .Cast<System.DirectoryServices.ActiveDirectory.DomainController>()
                                     .Select(dc => dc.Name)
                                     .Where(n => !string.IsNullOrWhiteSpace(n))
                                     .ToList();

                if (discovered.Count > 0)
                {
                    AppLog.Detail($"Discovered {discovered.Count} domain controller(s) from AD: {string.Join(", ", discovered)}");
                    return OrderPdcFirst(discovered);
                }
            }
            catch (Exception ex)
            {
                AppLog.Detail($"Domain controller discovery failed, falling back to configuration: {ex.Message}");
            }

            return OrderPdcFirst(_domainControllers);
        }// end of BuildDomainControllerSearchOrder

        /// <summary>Moves the PDC emulator to the front of the list, if it can be identified.</summary>
        private static List<string> OrderPdcFirst(List<string> domainControllers)
        {
            var ordered = new List<string>();
            try
            {
                string pdc = System.DirectoryServices.ActiveDirectory.Domain.GetComputerDomain().PdcRoleOwner.Name;
                if (!string.IsNullOrWhiteSpace(pdc)) ordered.Add(pdc);
            }
            catch
            {
                // Not domain-joined or the PDC can't be located; order doesn't matter then.
            }

            foreach (string dc in domainControllers)
            {
                // Compare short name against the PDC's FQDN so it isn't queried twice.
                bool alreadyQueued = ordered.Any(existing =>
                    existing.Equals(dc, StringComparison.OrdinalIgnoreCase) ||
                    existing.StartsWith(dc + ".", StringComparison.OrdinalIgnoreCase) ||
                    dc.StartsWith(existing + ".", StringComparison.OrdinalIgnoreCase));

                if (!alreadyQueued) ordered.Add(dc);
            }
            return ordered;
        }// end of OrderPdcFirst

        /// <summary>
        /// Short TCP probe of the WinRM port, so an offline DC costs well under a second instead of
        /// the remoting stack's own connect timeout.
        /// </summary>
        private static bool IsWinRmPortOpen(string host)
        {
            const int winRmPort = 5985;
            try
            {
                using var client = new System.Net.Sockets.TcpClient();
                return client.ConnectAsync(host, winRmPort).Wait(TimeSpan.FromMilliseconds(700)) && client.Connected;
            }
            catch
            {
                return false;
            }
        }// end of IsWinRmPortOpen

        // -----------------------------------------------------------------------------------------
        //                                        Search
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Finds users by partial name.
        ///
        /// Every other lookup in this tool demands an exact sAMAccountName, which is painful when
        /// all you have from a phone call is a surname.
        /// </summary>
        public void FindUsers(PrincipalContext context)
        {
            bool returnToMenu = false;
            do
            {
                ConsoleUi.PromptWithExit("Name, surname or username to search for");
                string term = ConsoleInput.ReadTrimmed();

                if (term.Equals("exit", StringComparison.OrdinalIgnoreCase)) { returnToMenu = true; continue; }
                if (term.Length < 2)
                {
                    ConsoleUi.Warn("Enter at least two characters.");
                    continue;
                }

                try
                {
                    // Three searches because PrincipalSearcher ANDs the properties set on the
                    // template, so one template cannot express "name OR surname OR account".
                    var matches = new SortedDictionary<string, UserPrincipal>(StringComparer.OrdinalIgnoreCase);
                    foreach (var template in new[]
                    {
                        new UserPrincipal(context) { Name = $"*{term}*" },
                        new UserPrincipal(context) { Surname = $"*{term}*" },
                        new UserPrincipal(context) { SamAccountName = $"*{term}*" }
                    })
                    {
                        using (template)
                        using (var searcher = new PrincipalSearcher(template))
                        using (var results = searcher.FindAll())
                        {
                            foreach (var found in results)
                            {
                                if (found is UserPrincipal u && !string.IsNullOrEmpty(u.SamAccountName)
                                    && !matches.ContainsKey(u.SamAccountName))
                                {
                                    matches[u.SamAccountName] = u;
                                }
                            }
                        }
                    }

                    ConsoleUi.Table(
                        new[] { "Username", "Display name", "Email", "Status" },
                        matches.Values.Select(u => new[]
                        {
                            u.SamAccountName,
                            u.DisplayName ?? "",
                            u.EmailAddress ?? "",
                            u.Enabled == true ? (u.IsAccountLockedOut() ? "LOCKED" : "Enabled") : "Disabled"
                        }));
                }
                catch (Exception ex)
                {
                    ConsoleUi.Fail($"Search for '{term}' failed: {ex.Message}", ex);
                }
            } while (!returnToMenu);
        }// end of FindUsers

        // -----------------------------------------------------------------------------------------
        //                                        Reports
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Disabled accounts whose "Delete on {date}" description has come due.
        ///
        /// Deactivation has always written that date into the description, but nothing ever read it
        /// back, so the 31-day cleanup depended on somebody remembering. This closes that loop.
        /// </summary>
        public void ReportAccountsDueForDeletion(PrincipalContext context)
        {
            ConsoleUi.Breadcrumb("Main", "Reports", "Accounts Due for Deletion");
            try
            {
                var rows = new List<(DateTime Due, string[] Cells)>();

                using (var template = new UserPrincipal(context) { Enabled = false })
                using (var searcher = new PrincipalSearcher(template))
                using (var results = searcher.FindAll())
                {
                    foreach (var result in results)
                    {
                        using (result)
                        {
                            if (!(result is UserPrincipal user)) continue;

                            DateTime? due = ParseDeletionDate(user.Description);
                            if (due == null) continue;

                            int daysLeft = (int)Math.Round((due.Value.Date - DateTime.Now.Date).TotalDays);
                            rows.Add((due.Value, new[]
                            {
                                user.SamAccountName ?? "",
                                user.DisplayName ?? "",
                                due.Value.ToString("yyyy-MM-dd"),
                                daysLeft <= 0 ? $"DUE ({-daysLeft}d ago)" : $"in {daysLeft}d"
                            }));
                        }
                    }
                }

                // Oldest first: whatever is most overdue needs attention first.
                ConsoleUi.Table(new[] { "Username", "Display name", "Delete on", "Status" },
                                rows.OrderBy(r => r.Due).Select(r => r.Cells));

                int overdue = rows.Count(r => r.Due.Date <= DateTime.Now.Date);
                if (overdue > 0) ConsoleUi.Warn($"{overdue} account(s) are at or past their deletion date.");
            }
            catch (Exception ex)
            {
                ConsoleUi.Fail($"Could not build the deletion report: {ex.Message}", ex);
            }
        }// end of ReportAccountsDueForDeletion

        /// <summary>
        /// Reads the date out of the "Delete on MM-dd-yyyy" description that deactivation writes.
        /// Tolerant of other separators and of extra text around it.
        /// </summary>
        private static DateTime? ParseDeletionDate(string description)
        {
            if (string.IsNullOrWhiteSpace(description)) return null;

            var match = System.Text.RegularExpressions.Regex.Match(
                description, @"delete\s*on\s*[: ]?\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (!match.Success) return null;

            // The writer uses MM-dd-yyyy.
            return int.TryParse(match.Groups[1].Value, out int month)
                && int.TryParse(match.Groups[2].Value, out int day)
                && int.TryParse(match.Groups[3].Value, out int year)
                && month >= 1 && month <= 12 && day >= 1 && day <= 31
                ? new DateTime(year, month, day)
                : (DateTime?)null;
        }// end of ParseDeletionDate

        /// <summary>
        /// Accounts whose password expires within a given number of days -- the calls this tool
        /// exists to field, caught before they turn into lockouts.
        /// </summary>
        public void ReportPasswordsExpiringSoon(PrincipalContext context)
        {
            ConsoleUi.Breadcrumb("Main", "Reports", "Passwords Expiring Soon");

            ConsoleUi.Prompt("Within how many days (Enter for 14)");
            string input = ConsoleInput.ReadTrimmed();
            int days = int.TryParse(input, out int parsed) && parsed > 0 ? parsed : 14;

            try
            {
                var rows = new List<(DateTime When, string[] Cells)>();

                using (var template = new UserPrincipal(context) { Enabled = true })
                using (var searcher = new PrincipalSearcher(template))
                using (var results = searcher.FindAll())
                {
                    foreach (var result in results)
                    {
                        using (result)
                        {
                            if (!(result is UserPrincipal user) || user.PasswordNeverExpires) continue;

                            DateTime expires = passwordManager.GetPasswordExpirationDate(user);
                            if (expires == DateTime.MinValue) continue;

                            int daysLeft = (int)Math.Round((expires.Date - DateTime.Now.Date).TotalDays);
                            if (daysLeft > days) continue;

                            rows.Add((expires, new[]
                            {
                                user.SamAccountName ?? "",
                                user.DisplayName ?? "",
                                expires.ToString("yyyy-MM-dd"),
                                daysLeft < 0 ? $"EXPIRED ({-daysLeft}d)" : daysLeft == 0 ? "TODAY" : $"{daysLeft}d"
                            }));
                        }
                    }
                }

                ConsoleUi.Table(new[] { "Username", "Display name", "Expires", "In" },
                                rows.OrderBy(r => r.When).Select(r => r.Cells));
            }
            catch (Exception ex)
            {
                ConsoleUi.Fail($"Could not build the expiry report: {ex.Message}", ex);
            }
        }// end of ReportPasswordsExpiringSoon

        /// <summary>
        /// Every account locked out recently, and the machine that did it.
        ///
        /// Reuses the WinRM Event 4740 sweep. Where "Find Lockout Source" answers "where was this
        /// one user locked out", this answers "what is going on" -- one workstation appearing
        /// against several accounts is the usual sign of a stale cached credential or a service
        /// running under an old password.
        /// </summary>
        public void ReportRecentLockouts()
        {
            ConsoleUi.Breadcrumb("Main", "Reports", "Recent Lockouts");

            ConsoleUi.Prompt("Look back how many hours (Enter for 24)");
            string input = ConsoleInput.ReadTrimmed();
            int hours = int.TryParse(input, out int parsed) && parsed > 0 ? parsed : 24;

            var events = QueryRecentLockoutEvents(hours);
            if (events == null) return;

            ConsoleUi.Table(new[] { "When", "Account", "Locked out from", "DC" },
                            events.OrderByDescending(e => e.When)
                                  .Select(e => new[] { e.When.ToString("MM-dd HH:mm:ss"), e.Account, e.Caller, e.DomainController }));

            var repeats = events.GroupBy(e => e.Caller, StringComparer.OrdinalIgnoreCase)
                                .Where(g => g.Select(e => e.Account).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1)
                                .ToList();
            foreach (var group in repeats)
            {
                ConsoleUi.Warn($"'{group.Key}' locked out {group.Select(e => e.Account).Distinct(StringComparer.OrdinalIgnoreCase).Count()} " +
                               "different accounts — check for a stale cached credential or a service account on that machine.");
            }
        }// end of ReportRecentLockouts

        /// <summary>
        /// Extracts the value of a named Data element from a Windows Event XML string.
        /// </summary>
        /// <param name="xml">The raw event XML.</param>
        /// <param name="dataName">The Name attribute to look for (e.g. "CallerComputerName").</param>
        /// <returns>The inner text value, or null if not found.</returns>
        private string ExtractXmlDataValue(string xml, string dataName)
        {
            try
            {
                string startTag = $"Name='{dataName}'>";
                int startIndex = xml.IndexOf(startTag, StringComparison.OrdinalIgnoreCase);
                if (startIndex == -1) return null;

                startIndex += startTag.Length;
                int endIndex = xml.IndexOf('<', startIndex);
                if (endIndex == -1) return null;

                return xml[startIndex..endIndex].Trim();
            }
            catch
            {
                return null;
            }
        }// end of ExtractXmlDataValue
    }// end of class
}// end of spacename
