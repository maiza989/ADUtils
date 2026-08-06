using Microsoft.Extensions.Configuration;
using Pastel;
using System.Diagnostics.Eventing.Reader;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.Security;

namespace ADUtils
{

    public class ActiveDirectoryManager
    {
        /// <summary>
        /// How long to wait on one DC's Security log before giving up. A firewall-blocked RPC call
        /// takes ~21 seconds to fail on its own, which is far longer than an operator should wait
        /// for an optional detail.
        /// </summary>
        private const int LockoutLookupTimeoutSeconds = 8;

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
        /// Set once the remote Security log is shown to be unreachable, so later lookups in the
        /// same session fail instantly instead of paying the timeout again.
        /// </summary>
        private bool _eventLogBlocked;

        public void SetAdminCredentials(string adminUsername, string adminPassword, IConfiguration configuration)
        {
            _adminUsername = adminUsername;
            _adminPassword = adminPassword;

            // myDomainName holds a *domain controller* hostname (e.g. "LMDC2"), which is what
            // PrincipalContext wants. EventLogSession needs the account's DOMAIN, so use myDomain.
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

                            AppLog.Screen($"\nFirst name: {user.GivenName ?? "N/A"}\n" +
                                              $"Last name: {user.Surname ?? "N/A"}\n" +
                                              $"Display name: {user.DisplayName ?? "N/A"}\n" +
                                              $"Username: {user.SamAccountName ?? "N/A"}\n" +
                                              $"Email: {user.EmailAddress ?? "N/A"}\n" +
                                              $"Title: {title}\n" +
                                              $"Department: {department}\n" +
                                              $"Member of: {userGroupsString}\n" +
                                              $"Password Last Set: {passwordManager.GetPasswordLastSetDate(user)}\n" +
                                              $"Password Expiration Date: {passwordManager.GetPasswordExpirationDate(user)}\n" +
                                              $"Bad Logon Counter: {user.BadLogonCount}\n" +
                                              $"Last Logon: {lastLogon}\n" +
                                              $"Last Bad Logon Attempt: {lastBadPwd}\n" +
                                              $"Account Status: {user.Enabled}\n" +
                                              $"Account Lockout Status: {user.IsAccountLockedOut()}\n" +
                                              $"Home Directory: {user.HomeDirectory ?? "N/A"}\n" +
                                              $"SID: {user.Sid}\n" +
                                              $"");
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
                AppLog.Warn("No domain controllers configured in Appsettings.json.", color: Color.DarkGoldenrod);
                return null;
            }

            SecureString securePassword = new SecureString();
            foreach (char c in _adminPassword)
                securePassword.AppendChar(c);
            securePassword.MakeReadOnly();

            // Filter server-side on both the event id and the target account, and cap the search to
            // a recent window. Without the time bound the DC scans the whole Security log, which on
            // a busy DC is far slower than the lookup is worth.
            long windowMs = (long)TimeSpan.FromDays(LockoutSearchWindowDays).TotalMilliseconds;
            string xpath = $"*[System[EventID=4740 and TimeCreated[timediff(@SystemTime) <= {windowMs}]]]" +
                           $" and *[EventData[Data[@Name='TargetUserName']='{username}']]";

            bool anyReachable = false;
            foreach (string dc in candidates)
            {
                if (!IsRpcPortOpen(dc))
                {
                    AppLog.Screen($"\t{dc}: not reachable — skipped.", Color.DarkGray);
                    continue;
                }

                anyReachable = true;
                AppLog.Prompt($"\t{dc}: ");

                string caller = null;
                string failure = null;
                bool rpcBlocked = false;
                DateTime? foundAt = null;

                // A blocked RPC call takes ~21 seconds to time out on its own, so bound the wait
                // instead of letting it dictate how long the operator sits there. The abandoned
                // thread unwinds on its own; nothing downstream depends on it.
                var probe = Task.Run(() =>
                {
                    try
                    {
                        using var session = new EventLogSession(dc, _domain, _adminUsername, securePassword, SessionAuthentication.Default);
                        var eventsQuery = new EventLogQuery("Security", PathType.LogName, xpath)
                        {
                            Session = session,
                            ReverseDirection = true     // newest first
                        };

                        using var logReader = new EventLogReader(eventsQuery);
                        using EventRecord evt = logReader.ReadEvent();

                        if (evt == null)
                        {
                            failure = $"no 4740 event for this user in the last {LockoutSearchWindowDays} days.";
                            return;
                        }

                        string found = ExtractXmlDataValue(evt.ToXml(), "CallerComputerName");
                        if (string.IsNullOrWhiteSpace(found))
                        {
                            failure = "event found but it carries no CallerComputerName.";
                            return;
                        }
                        // Don't write to the console from inside the task -- if the bounded wait
                        // below has already given up, output would interleave with the next prompt.
                        foundAt = evt.TimeCreated;
                        caller = found.TrimStart('\\');
                    }
                    catch (Exception ex) when (ex.HResult == unchecked((int)0x800706BA) ||
                                               ex.Message.Contains("RPC server is unavailable"))
                    {
                        // Port 135 answered but the RPC call itself failed, which means the dynamic
                        // RPC port range that 135 redirects to is blocked. No credential works
                        // around that, and it is a host firewall policy -- so it will be the same
                        // on every DC. Give up on the rest rather than paying the timeout again.
                        rpcBlocked = true;
                    }
                    catch (UnauthorizedAccessException)
                    {
                        failure = $"access denied for '{_adminUsername}' — the account needs to be in Event Log Readers on {dc}.";
                    }
                    catch (EventLogNotFoundException)
                    {
                        failure = "no Security log.";
                    }
                    catch (Exception ex)
                    {
                        failure = ex.Message;
                    }
                });

                if (!probe.Wait(TimeSpan.FromSeconds(LockoutLookupTimeoutSeconds)))
                {
                    AppLog.Warn($"gave up after {LockoutLookupTimeoutSeconds}s.", color: Color.DarkGoldenrod);
                    _eventLogBlocked = true;
                    PrintEventLogBlockedHelp();
                    return null;
                }

                if (caller != null)
                {
                    AppLog.Info($"found ({foundAt:MM-dd-yyyy HH:mm:ss}).", Color.DarkOliveGreen);
                    return caller;
                }
                if (rpcBlocked)
                {
                    AppLog.Warn("remote Security log is blocked.", color: Color.DarkGoldenrod);
                    _eventLogBlocked = true;
                    PrintEventLogBlockedHelp();
                    return null;
                }
                AppLog.Screen(failure, Color.DarkGray);
            }

            if (!anyReachable)
            {
                AppLog.Warn("\tNone of the configured domain controllers are reachable.", color: Color.IndianRed);
            }
            return null;
        }// end of LookupLockoutSource

        /// <summary>
        /// Explains the one thing that actually has to change for this lookup to work.
        /// </summary>
        private static void PrintEventLogBlockedHelp()
        {
            AppLog.Warn($"\tThe DCs' remote Security log is not readable from this machine.", color: Color.IndianRed);
            AppLog.Screen($"\tPort 135 answers but the RPC call is refused, so the dynamic RPC range is blocked.", Color.Gray);
            AppLog.Screen($"\tFix: enable the {"Remote Event Log Management".Pastel(Color.MediumPurple)} inbound firewall rules on the DCs,", Color.Gray);
            AppLog.Screen($"\tand ensure the admin account is in {"Event Log Readers".Pastel(Color.MediumPurple)}. Until then use ADUC / a DC session.", Color.Gray);
        }// end of PrintEventLogBlockedHelp

        /// <summary>
        /// Configured DCs, with the PDC emulator first because it is the DC that processes lockouts
        /// and so is the most likely to hold the 4740 record.
        /// </summary>
        private List<string> BuildDomainControllerSearchOrder()
        {
            var ordered = new List<string>();
            try
            {
                string pdc = System.DirectoryServices.ActiveDirectory.Domain.GetComputerDomain().PdcRoleOwner.Name;
                if (!string.IsNullOrWhiteSpace(pdc)) ordered.Add(pdc);
            }
            catch
            {
                // Not domain-joined or the PDC can't be located; fall back to the configured list.
            }

            foreach (string dc in _domainControllers)
            {
                // Match short name against the PDC's FQDN so it isn't probed twice.
                bool alreadyQueued = ordered.Any(existing =>
                    existing.Equals(dc, StringComparison.OrdinalIgnoreCase) ||
                    existing.StartsWith(dc + ".", StringComparison.OrdinalIgnoreCase));

                if (!alreadyQueued) ordered.Add(dc);
            }
            return ordered;
        }// end of BuildDomainControllerSearchOrder

        /// <summary>
        /// Short TCP probe of the RPC endpoint mapper, so an offline DC costs well under a second
        /// instead of the ~8 second RPC timeout.
        /// </summary>
        private static bool IsRpcPortOpen(string host)
        {
            try
            {
                using var client = new System.Net.Sockets.TcpClient();
                return client.ConnectAsync(host, 135).Wait(TimeSpan.FromMilliseconds(700)) && client.Connected;
            }
            catch
            {
                return false;
            }
        }// end of IsRpcPortOpen

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
