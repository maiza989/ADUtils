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
                AppLog.Prompt($"Enter the username to deactivate (type {"'exit'".Pastel(Color.MediumPurple)} to return to the main menu): ");
                string username = ConsoleInput.ReadTrimmedLower();

                if (username == "exit")
                {
                    returnToMenu = true;
                }// end of if statement
                else if (username.Length == 0)
                {
                    AppLog.Warn("Enter a username, or 'exit' to return to the menu.", color: Color.DarkGoldenrod);
                }
                else
                {
                    try
                    {
                        UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);              // Search for specific user using username
                        if (user == null)
                        {
                            AppLog.Warn($"\tUser account '{username}' not found in Active Directory.", color: Color.IndianRed);
                            continue;
                        }

                        // Check this BEFORE stripping groups. Previously the group removal ran
                        // unconditionally, so re-running on an already-disabled account silently
                        // stripped it again and still printed a success line.
                        if (user.Enabled != true)
                        {
                            AppLog.Warn($"User account '{username}' is {"ALREADY".Pastel(Color.MediumPurple)} disabled — no changes made.", color: Color.DarkGoldenrod);
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
                        AppLog.Info($"User account '{username}' removed from {removed} group(s); 'Domain Users' kept.", Color.LimeGreen);

                        user.Enabled = false;                                                                                   // Disabling the user account
                        user.Description = $"Delete on {deletionDateString}";                                                   // Change description with reminder of when to delete the ex user account
                        user.Save();
                        AppLog.Info($"User account '{username}' has been disabled\nAccount description changed to 'Delete on {deletionDateString}'", Color.LimeGreen);

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
                                AppLog.Info($"User account '{username}' has been moved to the {ACManager._myExEmployeeOU} OU", Color.LimeGreen);
                            }
                            catch (Exception ex)
                            {
                                AppLog.Error($"Move to '{ouPath}' FAILED: {ex.Message}", ex, Color.Crimson);
                                AppLog.Warn($"'{username}' is disabled but was NOT moved — move it manually.", color: Color.Crimson);
                            }
                        }// end of using

                        // Report what actually happened, not what was intended.
                        emailActionLog.Add(moved
                            ? $"User account '{username}' has been disabled and moved to the {ACManager._myExEmployeeOU} OU.\nAccount will be deleted on '{deletionDateString}'"
                            : $"User account '{username}' has been disabled but *** COULD NOT BE MOVED *** to the {ACManager._myExEmployeeOU} OU — move it manually.\nAccount will be deleted on '{deletionDateString}'");

                        // Disabling the account leaves the mailbox untouched, so mail to a departing
                        // employee simply goes unanswered. Offered rather than automatic because
                        // whether to auto-reply or hand the mailbox over is a per-departure decision.
                        HandleDepartingMailbox(username, user.DisplayName ?? username);
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

        /// <summary>
        /// Offers the mailbox side of an offboarding: hide from the address book, set an
        /// out-of-office reply, and hand access to a colleague.
        ///
        /// Each step is asked separately, and each is applied independently so one failure does not
        /// abandon the rest. Anything that fails is reported and added to the notification, never
        /// silently skipped.
        /// </summary>
        private void HandleDepartingMailbox(string username, string displayName)
        {
            if (!ConsoleUi.Confirm($"Also update the Exchange mailbox for '{username}'?"))
            {
                emailActionLog.Add($"Mailbox for '{username}' was left unchanged (skipped by operator).");
                return;
            }

            bool hide = ConsoleUi.Confirm("Hide the mailbox from the address book (GAL)?");
            bool autoReply = ConsoleUi.Confirm("Set an out-of-office auto-reply?");

            string autoReplyText = null;
            if (autoReply)
            {
                ConsoleUi.Prompt("Auto-reply message (Enter for a default)");
                autoReplyText = ConsoleInput.ReadTrimmed();
                if (autoReplyText.Length == 0)
                {
                    autoReplyText = $"{displayName} is no longer with the firm. " +
                                    "Please direct your message to the appropriate department and we will assist you.";
                }
            }

            string delegateTo = null;
            if (ConsoleUi.Confirm("Grant a colleague full access to this mailbox?"))
            {
                ConsoleUi.Prompt("Colleague's username");
                delegateTo = ConsoleInput.ReadTrimmed();
                if (delegateTo.Length == 0) delegateTo = null;
            }

            if (!hide && !autoReply && delegateTo == null)
            {
                ConsoleUi.Note("Nothing selected — mailbox left unchanged.");
                return;
            }

            try
            {
                using var exchange = new ExchangeSessionManager(Program.configuration);
                if (!exchange.Connect())
                {
                    ConsoleUi.Fail($"Mailbox for '{username}' was NOT updated — no Exchange session.");
                    emailActionLog.Add($"*** Mailbox for '{username}' could NOT be updated (no Exchange session) — handle it manually. ***");
                    return;
                }

                // Add/Remove-ADPermission and Set-Mailbox disagree with SMTP addresses, so resolve
                // once to a DN and use that throughout.
                if (!exchange.TryResolveMailbox(username, out string mailboxDn, out string mailboxName))
                {
                    emailActionLog.Add($"*** No mailbox found for '{username}' — nothing was changed in Exchange. ***");
                    return;
                }

                string dcParam = exchange.DomainController;

                if (hide)
                {
                    var p = new Dictionary<string, object>
                    {
                        ["Identity"] = mailboxDn,
                        ["HiddenFromAddressListsEnabled"] = true
                    };
                    if (dcParam != null) p["DomainController"] = dcParam;

                    if (exchange.RunCommand("Set-Mailbox", $"hiding '{mailboxName}' from the GAL", p))
                    {
                        ConsoleUi.Ok($"'{mailboxName}' hidden from the address book.");
                        emailActionLog.Add($"Mailbox '{mailboxName}' hidden from the global address list.");
                    }
                    else
                    {
                        emailActionLog.Add($"*** Could not hide mailbox '{mailboxName}' from the GAL — do it manually. ***");
                    }
                }

                if (autoReply)
                {
                    var p = new Dictionary<string, object>
                    {
                        ["Identity"] = mailboxDn,
                        ["AutoReplyState"] = "Enabled",
                        ["InternalMessage"] = autoReplyText,
                        ["ExternalMessage"] = autoReplyText
                    };
                    if (dcParam != null) p["DomainController"] = dcParam;

                    if (exchange.RunCommand("Set-MailboxAutoReplyConfiguration", $"setting the auto-reply for '{mailboxName}'", p))
                    {
                        ConsoleUi.Ok($"Out-of-office reply set on '{mailboxName}'.");
                        emailActionLog.Add($"Out-of-office auto-reply enabled on mailbox '{mailboxName}'.");
                    }
                    else
                    {
                        emailActionLog.Add($"*** Could not set the auto-reply on '{mailboxName}' — do it manually. ***");
                    }
                }

                if (delegateTo != null)
                {
                    var p = new Dictionary<string, object>
                    {
                        ["Identity"] = mailboxDn,
                        ["User"] = delegateTo,
                        ["AccessRights"] = "FullAccess",
                        ["InheritanceType"] = "All"
                    };
                    if (dcParam != null) p["DomainController"] = dcParam;

                    if (exchange.RunCommand("Add-MailboxPermission", $"granting '{delegateTo}' access to '{mailboxName}'", p))
                    {
                        ConsoleUi.Ok($"'{delegateTo}' granted full access to '{mailboxName}'.");
                        emailActionLog.Add($"'{delegateTo}' granted FullAccess to mailbox '{mailboxName}'.");
                    }
                    else
                    {
                        emailActionLog.Add($"*** Could not grant '{delegateTo}' access to '{mailboxName}' — do it manually. ***");
                    }
                }
            }
            catch (Exception ex)
            {
                ConsoleUi.Fail($"Error updating the mailbox for '{username}': {ex.Message}", ex);
                emailActionLog.Add($"*** Mailbox changes for '{username}' failed: {ex.Message} — handle it manually. ***");
            }
        }// end of HandleDepartingMailbox
    }// end of class
}// end of namespace
