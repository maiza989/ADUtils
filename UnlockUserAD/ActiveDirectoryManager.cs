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
        PasswordManager passwordManager = new PasswordManager();

        private string _adminUsername;
        private string _adminPassword;
        private string _domain;
        private List<string> _domainControllers = new List<string>();

        public void SetAdminCredentials(string adminUsername, string adminPassword, IConfiguration configuration)
        {
            _adminUsername = adminUsername;
            _adminPassword = adminPassword;
            _domain = configuration["AccountCreationSettings:myDomainName"];

            // Load all domain controllers from config — uses GetChildren() to avoid requiring Binder package
            _domainControllers = configuration.GetSection("AccountCreationSettings:myDomainControllers")
                                              .GetChildren()
                                              .Select(c => c.Value)
                                              .Where(v => !string.IsNullOrWhiteSpace(v))
                                              .ToList();

            if (_domainControllers.Count == 0)
                Console.WriteLine("Warning: No domain controllers configured in Appsettings.json.".Pastel(Color.DarkGoldenrod));
            else
                Console.WriteLine($"Loaded {_domainControllers.Count} domain controller(s): {string.Join(", ", _domainControllers)}".Pastel(Color.DarkCyan));
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
                Console.Write($"Enter the username to display info about (Type {"'exit'".Pastel(Color.MediumPurple)} to return to the main menu): ");
                string username = ConsoleInput.ReadTrimmedLower();

                if (username == "exit")
                {
                    returnToMenu = true;
                }// end of if statement
                else if (username.Length == 0)
                {
                    Console.WriteLine("Enter a username, or 'exit' to return to the main menu.".Pastel(Color.DarkGoldenrod));
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

                            Console.WriteLine($"\nFirst name: {user.GivenName ?? "N/A"}\n" +
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
                            Console.WriteLine($"\tUser '{username}' not found in Active Directory.".Pastel(Color.IndianRed));
                        }// end of else statement
                    }// end of try
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error looking up '{username}': {ex.Message}".Pastel(Color.IndianRed));
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
                Console.Write($"Enter the username to unlock (type {"'exit'".Pastel(Color.MediumPurple)} to return to the main menu): ");
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
                                Console.WriteLine($"\tUser account '{username}' has been unlocked.".Pastel(Color.LimeGreen));
                            }// end of inner-if-statement
                            else
                            {
                                Console.WriteLine($"\tUser account '{username}' is not locked.".Pastel(Color.OrangeRed));
                            }// end of else-statement
                        }// end of Outter-if-statement
                        else
                        {
                            Console.WriteLine($"\tUser account '{username}' not found in Active Directory.".Pastel(Color.IndianRed));
                        }// end of else-statement
                    }// end of Try-Catch
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error Unlocking a user: {ex.Message}".Pastel(Color.IndianRed));
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
                Console.WriteLine("\nUnlocking all user accounts...");
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
                                Console.WriteLine($"\t[{DateTime.Now:MM-dd-yyyy HH:mm:ss tt}]: User account '{user.SamAccountName}' has been unlocked.".Pastel(Color.LimeGreen));
                                anyUnlocked = true;
                            }
                            catch (Exception ex)
                            {
                                // Report and keep going -- one account we lack rights on shouldn't
                                // abandon the rest of the sweep.
                                Console.WriteLine($"\tCould not unlock '{user.SamAccountName}': {ex.Message}".Pastel(Color.IndianRed));
                            }
                        }
                    }// end of foreach
                }
                if (!anyUnlocked)                                                                                                              // If-Else statement to check if any user were unlocked and print appropriate response.
                {
                    Console.WriteLine("\tNo user accounts were locked.".Pastel(Color.DarkGoldenrod));
                }// end of if-statement
                else
                {
                    Console.WriteLine("\nAll user accounts have been unlocked successfully.".Pastel(Color.DarkCyan));
                }// end of else-statement
            }// end of Try-Catch
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}".Pastel(Color.IndianRed));
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
            Console.WriteLine("\nLocked user accounts:");
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
                    Console.WriteLine($"\tNo accounts are LOCKED!!! YAY!!!.".Pastel(Color.RoyalBlue));
                    return;
                }

                // One sweep of Event 4740 across the DCs for all locked accounts at once. Calling
                // this per user meant up to (users x DCs) connections and (users x DCs x 50) event
                // reads for a single menu selection.
                Dictionary<string, string> lockoutSources = GetLockoutSources();

                foreach (var locked in lockedUsers)
                {
                    string workstationName = lockoutSources.TryGetValue(locked.SamAccountName, out string caller) ? caller : "Unknown";
                    string when = locked.LockoutTime?.ToString("MM-dd-yyyy HH:mm:ss tt") ?? "time unavailable";

                    // Printed for every locked account. This line used to sit inside the
                    // "lockoutTime exists" branch, so accounts without a readable lockoutTime were
                    // counted as locked but never displayed.
                    Console.WriteLine($"\t[{when}] - {locked.SamAccountName} - Workstation: {workstationName}".Pastel(Color.Crimson));
                }// end of foreach
            }// end of try-catch
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}".Pastel(Color.IndianRed));
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
                Console.WriteLine($"Error reading lockoutTime for {user.SamAccountName}: {ex.Message}".Pastel(Color.DarkGoldenrod));
                return null;
            }
        }// end of ReadLockoutTime

        /// <summary>
        /// Sweeps every configured domain controller's Security log for Event ID 4740 (account
        /// locked out) and builds a map of sAMAccountName to the workstation that caused it.
        ///
        /// One sweep serves all locked accounts. Newest events win, so a user locked repeatedly
        /// reports the most recent source. Returns an empty map -- never throws -- if credentials
        /// or DCs are missing, so callers just display "Unknown".
        /// </summary>
        private Dictionary<string, string> GetLockoutSources()
        {
            const int maxEventsPerDc = 200;
            var sources = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrEmpty(_adminUsername) || string.IsNullOrEmpty(_adminPassword))
            {
                Console.WriteLine("Admin credentials not set — cannot determine lockout sources.".Pastel(Color.DarkGoldenrod));
                return sources;
            }

            if (_domainControllers.Count == 0)
            {
                Console.WriteLine("No domain controllers configured in Appsettings.json — cannot determine lockout sources.".Pastel(Color.DarkGoldenrod));
                return sources;
            }

            SecureString securePassword = new SecureString();
            foreach (char c in _adminPassword)
                securePassword.AppendChar(c);
            securePassword.MakeReadOnly();

            foreach (string dc in _domainControllers)
            {
                Console.WriteLine($"Checking {dc} for lockout events...".Pastel(Color.DarkCyan));
                try
                {
                    using var session = new EventLogSession(
                        dc,
                        _domain,
                        _adminUsername,
                        securePassword,
                        SessionAuthentication.Default);

                    EventLogQuery eventsQuery = new EventLogQuery("Security", PathType.LogName, "*[System[EventID=4740]]")
                    {
                        Session = session,
                        ReverseDirection = true     // newest first
                    };

                    using var logReader = new EventLogReader(eventsQuery);

                    int scanned = 0;
                    for (EventRecord evt = logReader.ReadEvent();
                         evt != null && scanned < maxEventsPerDc;
                         evt = logReader.ReadEvent(), scanned++)
                    {
                        using (evt)
                        {
                            string xml = evt.ToXml();

                            // Read the target name from its own field rather than testing whether the
                            // XML merely contains the username -- a substring test matched unrelated
                            // events for similarly named accounts (e.g. "jsmith" inside "jsmithers").
                            string target = ExtractXmlDataValue(xml, "TargetUserName");
                            if (string.IsNullOrWhiteSpace(target)) continue;

                            string callerName = ExtractXmlDataValue(xml, "CallerComputerName");
                            if (string.IsNullOrWhiteSpace(callerName)) continue;

                            // Newest first, so the first entry seen for a user is the latest one.
                            if (!sources.ContainsKey(target))
                            {
                                sources[target] = callerName.TrimStart('\\');
                            }
                        }
                    }
                }
                catch (Exception ex) when (ex.Message.Contains("RPC server is unavailable") ||
                                           ex.HResult == unchecked((int)0x800706BA))
                {
                    Console.WriteLine($"DC '{dc}' is unreachable (RPC unavailable) — skipping.".Pastel(Color.DarkGoldenrod));
                }
                catch (UnauthorizedAccessException)
                {
                    Console.WriteLine($"Access denied on '{dc}' — skipping.".Pastel(Color.DarkOrange));
                }
                catch (EventLogNotFoundException)
                {
                    Console.WriteLine($"Security log not found on '{dc}' — skipping.".Pastel(Color.DarkGoldenrod));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error querying '{dc}': {ex.Message} — skipping.".Pastel(Color.DarkOrange));
                }
            }

            return sources;
        }// end of GetLockoutSources

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
