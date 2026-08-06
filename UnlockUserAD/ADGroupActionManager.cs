using System.DirectoryServices.AccountManagement;
using Pastel;
using System.Drawing;

namespace ADUtils
{
    public class ADGroupActionManager
    {
        private const string EmailSubject = "ADUtil Action: Administrative Action in Active Directory";
        private const string MailboxEmailSubject = "ADUtil Action: Shared Mailbox Permission Changed";

        EmailNotifcationManager emailNotifcation = new EmailNotifcationManager(Program.configuration);
        AuditLogManager auditLogManager;
        List<string> emailActionLog = new List<string>();

        public ADGroupActionManager(AuditLogManager auditLogManager)
        {
            this.auditLogManager = auditLogManager;
        }

        /// <summary>
        /// A method to add user to group security and distrbuiton list in Active Directory.
        /// </summary>
        /// <param name="context"></param>
        public void AddUserToGroup(PrincipalContext context)
        {
            // Cleared per visit so re-entering the menu doesn't re-send earlier entries.
            emailActionLog.Clear();
            bool isExit = false;

            do
            {
                AppLog.Prompt($"Enter the username(Type {"'exit'".Pastel(Color.MediumPurple)} to go back to menu): ");
                string username = ConsoleInput.ReadTrimmed();
                if (username.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    isExit = true;
                    AppLog.Screen($"\nReturning to menu...");
                    break;
                }
                AppLog.Prompt($"Enter the group name (Type {"'exit'".Pastel(Color.MediumPurple)} to go back to menu): ");
                string groupName = ConsoleInput.ReadTrimmed();
                if (groupName.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    isExit = true;
                    AppLog.Screen($"\nReturning to menu...");
                    break;
                }
                else
                {
                    try
                    {
                        UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);                                              // Check for user in AD

                        if (user != null)
                        {
                            GroupPrincipal group = GroupPrincipal.FindByIdentity(context, groupName);                                                                   // Check for group in AD

                            if (group != null)
                            {
                                using (group)
                                {
                                    if (!group.Members.Contains(user))                                                                                                 // If the user is not in the group add him
                                    {
                                        group.Members.Add(user);                                                                                                       // Add the user to the group
                                        group.Save(context);                                                                                                           // Apply changes
                                        AppLog.Info($"User '{username}' added to group '{groupName}' successfully.", Color.LimeGreen);

                                        string logEntry = ($"\"{user.DisplayName}\" has been Added to \"{groupName}\" group in Active Directory\n");
                                        emailActionLog.Add(logEntry);
                                        auditLogManager.Log(logEntry);
                                    }// end of inner-2 if-statement
                                    else
                                    {
                                        AppLog.Warn($"User '{username}' is already a member of group '{groupName}'.", color: Color.DarkGoldenrod);
                                    }// end of inner-2 else-statement
                                }// end of using
                            }// end of inner if-statement
                            else
                            {
                                AppLog.Warn($"Group '{groupName}' not found in Active Directory.", color: Color.IndianRed);
                            }// end of outter else-statement
                        }// end of outter if-statement
                        else
                        {
                            AppLog.Warn($"User '{username}' not found in Active Directory.", color: Color.IndianRed);
                        }
                    }// end of try
                    catch (Exception ex)
                    {
                        AppLog.Error($"Error adding user to group: {ex.Message}", ex, Color.IndianRed);
                    }// end of catch
                }
            } while (!isExit);

            SendActionLog(EmailSubject);
        }// end of AddUserToGroup

        /// <summary>
        /// A method that remove a user from a security group and distrubtion list in Active Directory
        /// </summary>
        /// <param name="context"></param>
        public void RemoveUserFromGroup(PrincipalContext context)
        {
            emailActionLog.Clear();
            bool isExit = false;

            do
            {
                AppLog.Prompt($"Enter the username(Type {"'exit'".Pastel(Color.MediumPurple)} to go back to menu): ");
                string username = ConsoleInput.ReadTrimmed();
                if (username.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    isExit = true;
                    AppLog.Screen($"\nReturning to menu...");
                    break;
                }
                AppLog.Prompt($"Enter the group name (Type {"'exit'".Pastel(Color.MediumPurple)} to go back to menu): ");
                string groupName = ConsoleInput.ReadTrimmed();
                if (groupName.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    isExit = true;
                    AppLog.Screen($"\nReturning to menu...");
                    break;
                }
                else
                {
                    try
                    {
                        UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);                                             // Check for user in AD

                        if (user != null)
                        {
                            GroupPrincipal group = GroupPrincipal.FindByIdentity(context, groupName);                                                                  // Check for group in AD

                            if (group != null)
                            {
                                using (group)
                                {
                                    if (group.Members.Contains(user))
                                    {
                                        group.Members.Remove(user);
                                        group.Save();                                                                                                                 // Apply changes
                                        AppLog.Info($"User '{username}' removed from group '{groupName}' successfully.", Color.LimeGreen);

                                        string logEntry = ($"\"{user.DisplayName}\" has been removed from \"{groupName}\" group in Active Directory\n");
                                        emailActionLog.Add(logEntry);
                                        auditLogManager.Log(logEntry);
                                    }// end of inner-2 if-statement
                                    else
                                    {
                                        AppLog.Warn($"User '{username}' is not a member of group '{groupName}'.", color: Color.DarkGoldenrod);
                                    }// end of inner-2 else-statement
                                }// end of using
                            }// end of inner if-statement
                            else
                            {
                                AppLog.Warn($"Group '{groupName}' not found in Active Directory.", color: Color.IndianRed);
                            }// end of outter else-statement
                        }// end of outter if-statement
                        else
                        {
                            AppLog.Warn($"User '{username}' not found in Active Directory.", color: Color.IndianRed);
                        }
                    }// end of try
                    catch (Exception ex)
                    {
                        AppLog.Error($"Error removing user from group: {ex.Message}", ex, Color.IndianRed);
                    }// end of catch
                }
            } while (!isExit);

            SendActionLog(EmailSubject);
        }// end of RemoveUserFromGroup

        /// <summary>
        /// Grants a user FullAccess and Send As on an on-prem Exchange shared mailbox.
        /// </summary>
        public void AddUserToSharedMailbox(PrincipalContext context)
        {
            ChangeSharedMailboxAccess(context, granting: true);
        }// end of AddUserToSharedMailbox

        /// <summary>
        /// Revokes a user's FullAccess and Send As on an on-prem Exchange shared mailbox.
        /// </summary>
        public void RemoveUserFromSharedMailbox(PrincipalContext context)
        {
            ChangeSharedMailboxAccess(context, granting: false);
        }// end of RemoveUserFromSharedMailbox

        /// <summary>
        /// Shared implementation for granting and revoking shared-mailbox access.
        ///
        /// Uses the on-prem Exchange cmdlet pairs -- Add/Remove-MailboxPermission for FullAccess and
        /// Add/Remove-ADPermission for Send As. The Exchange Online equivalent
        /// (Add-RecipientPermission via Connect-ExchangeOnline) does not exist on Exchange 2016,
        /// which is where this deployment's mailboxes live.
        /// </summary>
        private void ChangeSharedMailboxAccess(PrincipalContext context, bool granting)
        {
            emailActionLog.Clear();
            string verb = granting ? "grant" : "revoke";
            bool isExit = false;

            do
            {
                AppLog.Prompt($"Enter the username to {verb} access for (Type {"'exit'".Pastel(Color.MediumPurple)} to go back to menu): ");
                string username = ConsoleInput.ReadTrimmed();
                if (username.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    isExit = true;
                    AppLog.Screen("\nReturning to menu...");
                    break;
                }// end of if statement

                AppLog.Prompt($"Enter the shared mailbox email or alias (Type {"'exit'".Pastel(Color.MediumPurple)} to go back to menu): ");
                string sharedMailbox = ConsoleInput.ReadTrimmed();
                if (sharedMailbox.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    isExit = true;
                    AppLog.Screen("\nReturning to menu...");
                    break;
                }// end of if statement

                if (username.Length == 0 || sharedMailbox.Length == 0)
                {
                    AppLog.Warn("Both a username and a shared mailbox are required.", color: Color.DarkGoldenrod);
                    continue;
                }

                try
                {
                    // Confirm the account exists in AD before asking Exchange to do anything, so a
                    // typo produces a clear message instead of an opaque Exchange error.
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
                    if (user == null)
                    {
                        AppLog.Warn($"User '{username}' not found in Active Directory.", color: Color.IndianRed);
                        continue;
                    }

                    using (var exchange = new ExchangeSessionManager(Program.configuration))
                    {
                        if (!exchange.Connect())
                        {
                            AppLog.Warn($"Could not {verb} access — no Exchange session.", color: Color.IndianRed);
                            continue;
                        }

                        // FullAccess -- lets the user open the mailbox.
                        var mailboxParams = new Dictionary<string, object>
                        {
                            ["Identity"] = sharedMailbox,
                            ["User"] = username,
                            ["AccessRights"] = "FullAccess"
                        };
                        if (granting)
                        {
                            mailboxParams["InheritanceType"] = "All";
                        }
                        else
                        {
                            mailboxParams["Confirm"] = false;
                        }
                        if (exchange.DomainController != null)
                        {
                            mailboxParams["DomainController"] = exchange.DomainController;
                        }

                        bool fullAccessOk = exchange.RunCommand(
                            granting ? "Add-MailboxPermission" : "Remove-MailboxPermission",
                            $"{(granting ? "granting" : "revoking")} FullAccess on '{sharedMailbox}' for '{username}'",
                            mailboxParams);

                        // Send As -- on-prem uses an AD extended right, not Add-RecipientPermission.
                        var sendAsParams = new Dictionary<string, object>
                        {
                            ["Identity"] = sharedMailbox,
                            ["User"] = username,
                            ["ExtendedRights"] = "Send As",
                            ["Confirm"] = false
                        };
                        if (exchange.DomainController != null)
                        {
                            sendAsParams["DomainController"] = exchange.DomainController;
                        }

                        bool sendAsOk = exchange.RunCommand(
                            granting ? "Add-ADPermission" : "Remove-ADPermission",
                            $"{(granting ? "granting" : "revoking")} Send As on '{sharedMailbox}' for '{username}'",
                            sendAsParams);

                        if (!fullAccessOk && !sendAsOk)
                        {
                            AppLog.Warn($"Nothing was changed for '{username}' on '{sharedMailbox}'.", color: Color.IndianRed);
                            continue;
                        }

                        // Log only the rights that actually changed.
                        var changed = new List<string>();
                        if (fullAccessOk) changed.Add("FullAccess");
                        if (sendAsOk) changed.Add("Send As");

                        string direction = granting ? "granted" : "revoked";
                        string preposition = granting ? "on" : "from";
                        AppLog.Info($"{string.Join(" and ", changed)} {direction} for '{username}' {preposition} '{sharedMailbox}'.", Color.LimeGreen);

                        if (!fullAccessOk || !sendAsOk)
                        {
                            string missing = fullAccessOk ? "Send As" : "FullAccess";
                            AppLog.Warn($"{missing} was NOT {direction} — apply it manually in Exchange.", color: Color.DarkGoldenrod);
                        }

                        string logEntry = $"\"{username}\" — {string.Join(" and ", changed)} {direction} {preposition} shared mailbox \"{sharedMailbox}\" in Exchange\n";
                        emailActionLog.Add(logEntry);
                        auditLogManager.Log(logEntry);
                    }// end of using
                }// end of try
                catch (Exception ex)
                {
                    AppLog.Error($"Error changing shared mailbox access: {ex.Message}", ex, Color.IndianRed);
                }// end of catch
            } while (!isExit);

            SendActionLog(MailboxEmailSubject);
        }// end of ChangeSharedMailboxAccess

        /// <summary>
        /// A method that list all group secuirty and distrubtion list in Active Directory.
        /// </summary>
        /// <param name="context"></param>
        public void ListAllGroups(PrincipalContext context)
        {
            AppLog.Screen("\nList of all groups:");

            try
            {
                AppLog.Prompt($"Enter the {"first letter".Pastel(Color.MediumPurple)} of the group name to filter by (or press Enter to show all groups): ");
                char filterLetter = Console.ReadKey().KeyChar;
                AppLog.Blank();

                List<string> groupNames = new List<string>();
                using (PrincipalSearcher searcher = new PrincipalSearcher(new GroupPrincipal(context)))                                                               // Search for all groups
                using (var results = searcher.FindAll())
                {
                    foreach (var result in results)
                    {
                        using (result)
                        {
                            string name = (result as GroupPrincipal)?.Name;
                            if (string.IsNullOrEmpty(name)) continue;                                                                                                 // Guard: indexing name[0] on an empty name threw

                            if (char.ToLower(name[0]) == char.ToLower(filterLetter) || filterLetter == '\r')                                                          // Filter by the first letter or show all groups if Enter is pressed)
                            {
                                groupNames.Add(name);
                            }
                        }// end of using
                    }// end of foreach
                }// end of using

                if (groupNames.Count == 0)
                {
                    // Previously Max() on an empty list threw "Sequence contains no elements" here.
                    AppLog.Warn(filterLetter == '\r'
                        ? "No groups found in Active Directory."
                        : $"No groups start with '{filterLetter}'.");
                    return;
                }

                groupNames.Sort();
                PrintInColumns(groupNames);
            }// end of try
            catch (Exception ex)
            {
                AppLog.Error($"Error listing groups: {ex.Message}", ex, Color.IndianRed);
            }// end of catch
        }// end of ListAllGroups

        /// <summary>
        /// A method that searches for the members of a specified group in Active Directory and lists them in a grid style.
        /// </summary>
        /// <param name="context"></param>
        public void ListGroupMembers(PrincipalContext context)
        {
            bool isExit = false;
            do
            {
                AppLog.Prompt($"Enter the group name (Type {"'exit'".Pastel(Color.MediumPurple)} to go back to menu): ");
                string groupName = ConsoleInput.ReadTrimmed();
                if (groupName.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    isExit = true;
                    AppLog.Screen("\nReturning to menu...");
                    break;
                }
                else
                {
                    try
                    {
                        GroupPrincipal group = GroupPrincipal.FindByIdentity(context, groupName);                                                          // Check for group in AD

                        if (group != null)
                        {
                            using (group)
                            {
                                AppLog.Screen($"\nMembers of group '{groupName}':");

                                List<string> memberNames = new List<string>();
                                foreach (var member in group.GetMembers())
                                {
                                    if (!string.IsNullOrEmpty(member.SamAccountName))
                                    {
                                        memberNames.Add(member.SamAccountName);
                                    }
                                }// end of foreach

                                if (memberNames.Count > 0)
                                {
                                    memberNames.Sort();
                                    PrintInColumns(memberNames);
                                }// end of if-statement
                                else
                                {
                                    AppLog.Screen("No members found in this group.");
                                }// end of else-statement
                            }// end of using
                        }// end of if-statements
                        else
                        {
                            AppLog.Warn($"Group '{groupName}' not found in Active Directory.", color: Color.IndianRed);
                        }// end of else
                    }// end of try-catch
                    catch (Exception ex)
                    {
                        AppLog.Error($"Error listing members of group: {ex.Message}", ex, Color.IndianRed);
                    }// end of catch
                }// end of else
            } while (!isExit);
        }// end of ListGroupMembers

        /// <summary>
        /// Prints names in a column grid sized to the console width.
        ///
        /// Shared by ListAllGroups and ListGroupMembers, which previously carried identical copies.
        /// Both divided by a column count that could be zero when a name was wider than the window;
        /// (int)+Infinity is int.MinValue, so the loop never ran and nothing printed at all.
        /// </summary>
        private static void PrintInColumns(List<string> names)
        {
            int columnWidth = names.Max(n => n.Length) + 5;                                                                                                          // Add padding

            int windowWidth;
            try
            {
                windowWidth = Console.WindowWidth;
            }
            catch (IOException)
            {
                windowWidth = 80;                                                                                                                                    // Redirected output has no window
            }

            int numColumns = Math.Max(1, windowWidth / columnWidth);                                                                                                 // Never 0 — see remarks
            int numRows = (int)Math.Ceiling((double)names.Count / numColumns);

            // Build each row then emit it in one go, so the log records one line per row rather
            // than one entry per cell.
            for (int i = 0; i < numRows; i++)                                                                                                                        // Nested for loop to print names in a grid style
            {
                var row = new System.Text.StringBuilder();
                for (int j = 0; j < numColumns; j++)
                {
                    int index = i + j * numRows;                                                                                                                     // Calculate the index based on 'i' rows and 'j' columns
                    if (index < names.Count)
                    {
                        row.Append($"- {names[index].PadRight(columnWidth)}");                                                                                        // Each name with specified right padding
                    }
                }// end of inner for loop
                AppLog.Screen(row.ToString().TrimEnd());
            }// end of outter for loop
        }// end of PrintInColumns

        /// <summary>
        /// Sends the accumulated action log, if any, and clears it so the next visit starts clean.
        /// </summary>
        private void SendActionLog(string subject)
        {
            if (emailActionLog.Count == 0) return;

            string emailBody = string.Join("\n", emailActionLog);
            emailNotifcation.SendEmailNotification(subject, emailBody);
            emailActionLog.Clear();
        }// end of SendActionLog
    }// end of class
}// end of namespace
