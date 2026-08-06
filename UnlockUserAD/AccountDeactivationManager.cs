using Pastel;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Drawing;


namespace ADUtils
{
    public class AccountDeactivationManager
    {
        private List<string> emailActionLog = new List<string>();
        EmailNotifcationManager emailNotification = new EmailNotifcationManager(Program.configuration);

        public void DeactivateUserAccount(PrincipalContext context, string adminUsername, string adminPassword)
        {
            // Start a fresh log each visit, otherwise re-entering the menu re-sends every
            // deactivation from earlier in the session.
            emailActionLog.Clear();

            AccountCreationManager ACManager = new AccountCreationManager(Program.configuration);

            // The second DC component is the *other* half of the domain. This used to interpolate
            // _myDomain twice, producing "DC=lmlawfirm,DC=lmlawfirm", so the bind always failed and
            // no account was ever actually moved -- while the notification email reported success.
            string ouPath = $"LDAP://OU={ACManager._myExEmployeeOU},OU={ACManager._myConfiguredParentOU},DC={ACManager._myDomain},DC={ACManager._myDomainDotCom}";
            DateTime deletionDate = DateTime.Now.AddDays(ACManager._myDeletionGraceDays);
            string deletionDateString = deletionDate.ToString("MM-dd-yyyy");
            bool returnToMenu = false;

            do
            {
                Console.Write($"Enter the username to deactivate (type {"'exit'".Pastel(Color.MediumPurple)} to return to the main menu): ");
                string username = ConsoleInput.ReadTrimmedLower();

                if (username == "exit")
                {
                    returnToMenu = true;
                }// end of if statement
                else if (username.Length == 0)
                {
                    Console.WriteLine("Enter a username, or 'exit' to return to the menu.".Pastel(Color.DarkGoldenrod));
                }
                else
                {
                    try
                    {
                        UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);              // Search for specific user using username
                        if (user == null)
                        {
                            Console.WriteLine($"\tUser account '{username}' not found in Active Directory.".Pastel(Color.IndianRed));
                            continue;
                        }

                        // Check this BEFORE stripping groups. Previously the group removal ran
                        // unconditionally, so re-running on an already-disabled account silently
                        // stripped it again and still printed a success line.
                        if (user.Enabled != true)
                        {
                            Console.WriteLine($"User account '{username}' is {"ALREADY".Pastel(Color.MediumPurple)} disabled — no changes made.".Pastel(Color.DarkGoldenrod));
                            continue;
                        }

                        // Materialize before mutating: Members.Remove/Save while enumerating the
                        // live GetGroups() result mutates the collection being iterated.
                        var groups = user.GetGroups().OfType<GroupPrincipal>().ToList();
                        int removed = 0;
                        foreach (GroupPrincipal group in groups)
                        {
                            if (group.Name == "Domain Users") continue;                                                                // Keep the primary group

                            try
                            {
                                group.Members.Remove(user);
                                group.Save();
                                removed++;
                            }
                            catch (Exception ex)
                            {
                                AppLog.Warn($"Could not remove '{username}' from '{group.Name}': {ex.Message}", ex, Color.DarkGoldenrod);
                            }
                        }// end of foreach
                        Console.WriteLine($"User account '{username}' removed from {removed} group(s); 'Domain Users' kept.".Pastel(Color.LimeGreen));

                        user.Enabled = false;                                                                                   // Disabling the user account
                        user.Description = $"Delete on {deletionDateString}";                                                   // Change description with reminder of when to delete the ex user account
                        user.Save();
                        Console.WriteLine($"User account '{username}' has been disabled\nAccount description changed to 'Delete on {deletionDateString}'".Pastel(Color.LimeGreen));

                        // GetUnderlyingObject() is owned by the UserPrincipal -- don't dispose it.
                        DirectoryEntry userEntry = (DirectoryEntry)user.GetUnderlyingObject();
                        userEntry.CommitChanges();

                        bool moved = false;
                        using (DirectoryEntry startOU = new DirectoryEntry(userEntry.Path))
                        using (DirectoryEntry endOU = new DirectoryEntry(ouPath, adminUsername, adminPassword))
                        {
                            try
                            {
                                startOU.MoveTo(endOU);
                                moved = true;
                                Console.WriteLine($"User account '{username}' has been moved to the {ACManager._myExEmployeeOU} OU".Pastel(Color.LimeGreen));
                            }
                            catch (Exception ex)
                            {
                                AppLog.Error($"Move to '{ouPath}' FAILED: {ex.Message}", ex, Color.Crimson);
                                Console.WriteLine($"'{username}' is disabled but was NOT moved — move it manually.".Pastel(Color.Crimson));
                            }
                        }// end of using

                        // Report what actually happened, not what was intended.
                        emailActionLog.Add(moved
                            ? $"User account '{username}' has been disabled and moved to the {ACManager._myExEmployeeOU} OU.\nAccount will be deleted on '{deletionDateString}'"
                            : $"User account '{username}' has been disabled but *** COULD NOT BE MOVED *** to the {ACManager._myExEmployeeOU} OU — move it manually.\nAccount will be deleted on '{deletionDateString}'");
                    }// end of try
                    catch (Exception ex)
                    {
                        AppLog.Error($"Error deactivating user '{username}': {ex.Message}", ex, Color.Crimson);
                    }// end of catch
                }// end of else statement
            } while (!returnToMenu);

            if (emailActionLog.Count > 0)
            {
                string emailBody = string.Join("\n", emailActionLog);
                emailNotification.SendEmailNotification("ADUtil Action: Administrative Action in Active Directory", emailBody);
            }// end of if statement
        }// end of DeactivateUserAccount
    }// end of class
}// end of namespace
