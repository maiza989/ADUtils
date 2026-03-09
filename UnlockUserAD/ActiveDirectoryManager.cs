
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using ADUtils;
using System.Diagnostics;
using System.Linq;
using Pastel;
using System.Drawing;
using System.Diagnostics.Eventing.Reader;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Security;

namespace ADUtils
{

    public class ActiveDirectoryManager
    {
        PasswordManager passwordManager = new PasswordManager();
        /*
                // ✅ ADDED: P/Invoke for LogonUser
                [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
                private static extern bool LogonUser(
                    string username,
                    string domain,
                    string password,
                    int logonType,
                    int logonProvider,
                    out IntPtr token);

                [DllImport("kernel32.dll", SetLastError = true)]
                private static extern bool CloseHandle(IntPtr handle);

                private const int LOGON32_LOGON_INTERACTIVE = 2;
                private const int LOGON32_PROVIDER_DEFAULT = 0;
                const int LOGON32_LOGON_NEW_CREDENTIALS = 9;

                // ✅ ADDED: Admin credential properties
                private readonly string _domain = "lmdc4";
                private readonly string _adminUsername = "admmalghamgham";
                private readonly string _adminPassword = "****"; // TODO - Update this when testing
        */
        /*   public ActiveDirectoryManager(string domain, string adminUsername, string adminPassword)
           {
               _domain = domain;
               _adminUsername = adminUsername;
               _adminPassword = adminPassword;
           }*/

        /*    // ✅ ADDED: Helper to run any AD operation as the admin user
            private void RunAsAdmin(Action action)
            {
                IntPtr adminToken = IntPtr.Zero;
                try
                {
                    bool success = LogonUser(
                        _adminUsername,
                        _domain,
                        _adminPassword,
                        LOGON32_LOGON_NEW_CREDENTIALS,
                        LOGON32_PROVIDER_DEFAULT,
                        out adminToken);

                    if (!success)
                        throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());

                    using (WindowsIdentity adminIdentity = new WindowsIdentity(adminToken))
                    {
                        WindowsIdentity.RunImpersonated(adminIdentity.AccessToken, action);
                    }
                }
                finally
                {
                    if (adminToken != IntPtr.Zero)
                        CloseHandle(adminToken);
                }
            }
    */
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
                string username = Console.ReadLine().Trim().ToLower();

                if (username.ToLower().Trim() == "exit")
                {
                    returnToMenu = true;
                }// end of if statement
                else
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
                    if (user != null)
                    {
                        var groups = user.GetGroups();
                        foreach (var group in groups)
                        {
                            userGroups.Add(group.Name);
                        }// end of foreach
                        string userGroupsString = string.Join(", ", userGroups);

                        DirectoryEntry directoryEntry = user.GetUnderlyingObject() as DirectoryEntry;
                        string title = directoryEntry.Properties["title"].Value as string;
                        string department = directoryEntry.Properties["department"].Value as string;

                        // TODO - Make a check if the password or last logon date are null                        
                        DateTime lastBadPasswordAttemptLocal = TimeZoneInfo.ConvertTimeFromUtc(user.LastBadPasswordAttempt.Value.ToUniversalTime(), TimeZoneInfo.Local);
                        DateTime lastLogonLocal = TimeZoneInfo.ConvertTimeFromUtc(user.LastLogon.Value.ToUniversalTime(), TimeZoneInfo.Local);

                        Console.WriteLine($"\nFirst name: {user.GivenName ?? "N/A"}\n" +
                                          $"Last name: {user.Surname ?? "N/A"}\n" +
                                          $"Display name: {user.DisplayName ?? "N/A"}\n" +
                                          $"Username: {user.SamAccountName ?? "N/A"}\n" +
                                          $"Email: {user.EmailAddress ?? "N/A"}\n" +
                                          $"Title: {title ?? "N/A"}\n" +
                                          $"Department: {department ?? "N/A"}\n" +
                                          $"Member of: {userGroupsString ?? "N/A"}\n" +
                                          $"Password Last Set: {passwordManager.GetPasswordLastSetDate(user)}\n" +
                                          $"Password Experation Date: {passwordManager.GetPasswordExpirationDate(user)}\n" +
                                          $"Bad Logon Counter: {user.BadLogonCount}\n" +
                                          $"Last Logon: {lastLogonLocal}\n" +
                                          $"Last Bad Logon Attempt: {lastBadPasswordAttemptLocal}\n" +
                                          $"Account Status: {user.Enabled}\n" +
                                          $"Account Lockout Status: {user.IsAccountLockedOut()}\n" +
                                          $"Home Directory: {user.HomeDirectory ?? "N/A"}\n" +
                                          $"SID: {user.Sid}\n" +
                                          $"");

                    }// end of if statement
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
                string username = Console.ReadLine().Trim().ToLower();

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
                PrincipalSearcher searcher = new PrincipalSearcher(new UserPrincipal(context) { Enabled = true });
                bool anyUnlocked = false;
                foreach (var result in searcher.FindAll())
                {
                    UserPrincipal user = result as UserPrincipal;
                    if (!user.IsAccountLockedOut())
                    {
                        continue;
                    }
                    if (user != null && user.IsAccountLockedOut())                                                                             // If-statement to unlock all users
                    {
                        user.UnlockAccount();
                        Console.WriteLine($"\t[{DateTime.Now:MM-dd-yyyy HH:mm:ss tt}]: User account '{user.SamAccountName}' has been unlocked.".Pastel(Color.LimeGreen));
                        anyUnlocked = true;
                    }// end of if-statement
                }// end of foreach
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
            // RunAsAdmin(() =>
            // {
            Console.WriteLine("\nLocked user accounts:");
            try
            {
                PrincipalSearcher searcher = new PrincipalSearcher(new UserPrincipal(context) { Enabled = true });                              // Creating the search object
                bool isAnyLocked = false;
                foreach (var result in searcher.FindAll())                                                                                      // Look through what is in the user search object
                {
                    UserPrincipal user = result as UserPrincipal;
                    if (!user.IsAccountLockedOut() || user == null)
                    {
                        continue;
                    }
                    if (user != null && user.IsAccountLockedOut())                                                                              // Print out all locked users
                    {
                        DirectoryEntry directoryEntry = (user.GetUnderlyingObject() as DirectoryEntry);
                        DateTime? lockoutTime = null;
                        string workstationName = "N/A";


                        // TODO - DONE Fix grabbing time lock out for users.
                        if (directoryEntry.Properties.Contains("lockoutTime"))
                        {
                            object lockOutValue = directoryEntry.Properties["lockoutTime"].Value;
                            if (lockOutValue != null)
                            {
                                long lockoutTicks = 0;
                                try
                                {
                                    var highPart = (int)lockOutValue.GetType().InvokeMember("HighPart", System.Reflection.BindingFlags.GetProperty, null, lockOutValue, null);
                                    var lowPart = (int)lockOutValue.GetType().InvokeMember("LowPart", System.Reflection.BindingFlags.GetProperty, null, lockOutValue, null);
                                    lockoutTicks = ((long)highPart << 32) + (uint)lowPart;
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error reading lockoutTime COM object for {user.SamAccountName}: {ex.Message}");
                                }
                                if (lockoutTicks > 0)
                                {
                                    lockoutTime = DateTime.FromFileTimeUtc(lockoutTicks).ToLocalTime();
                                }
                            }
                            //    workstationName = GetWorkstationNameFromEvent(user.SamAccountName);
                            Console.WriteLine($"\t[{lockoutTime?.ToString("MM-dd-yyyy HH:mm:ss tt")}] - {user.SamAccountName} - Workstation: {workstationName}".Pastel(Color.Crimson));
                        }
                        isAnyLocked = true;
                    }// end of if-statement
                }// end of foreach
                if (!isAnyLocked)
                {
                    Console.WriteLine($"\tNo accounts are LOCKED!!! YAY!!!.".Pastel(Color.RoyalBlue));
                }// end of if-statement
            }// end of try-catch
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}".Pastel(Color.IndianRed));
            }// end of catch
             //  });
        }// end of CheckLockedAccounts
    }// end of class
}// end of spacename
