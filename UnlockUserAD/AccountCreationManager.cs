using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Task = System.Threading.Tasks.Task;
using System.Diagnostics;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using Microsoft.VisualBasic;
using System.Collections.ObjectModel;


// TODO - User Account Creation: Enable users to create new accounts in Active Directory. 
namespace UnlockUserAD
{
    public class AccountCreationManager
    {
        string myDomain = Environment.GetEnvironmentVariable("MY_DOMAIN");
        string mydomainDotCom = Environment.GetEnvironmentVariable("MY_DOMAIN.COM");
        string myParentOU = Environment.GetEnvironmentVariable("MY_PARENT_OU");
        string myCompany = Environment.GetEnvironmentVariable("MY_COMPANY");
        string myExhcangeDatabase = Environment.GetEnvironmentVariable("MY_EXCHANGE_DATABASE");

        string firstName;
        string lastName;
        string jobTitle;
        string departmentEntry;
        string description;
        string office;
        string manager;
        string targetOU;
        string firstInitial;
        string lastInitial;
        string username;
        string email;
        string password;
        string userProfile;
        string clsUserFolder;

        public AccountCreationManager() 
        {
    
        }

        public void CreateUserAccount(string adminUsername, string adminPassword)
        {
            // Prompt for user details
            Console.Write("Enter new user's first name: ");
            firstName = Console.ReadLine();

            Console.Write("Enter new user's last name: ");
            lastName = Console.ReadLine();

            Console.Write("Enter user's job title: ");
            jobTitle = Console.ReadLine();

            Console.Write("Enter user's department: ");
            departmentEntry = Console.ReadLine();

            Console.Write("Enter user description: ");
            description = Console.ReadLine();

            Console.Write("Enter user office (KY, MI, GA): ");
            office = Console.ReadLine();

            /*Console.Write("Enter user manager (SAM Account Name): ");
            manager = Console.ReadLine();*/

            Console.WriteLine("Select New User OU (Admin Staff, Collector, Atty, Acct, IT, Compliance, Michigan_Users, Cooling_Users)");
            Console.Write("Enter your choice:");
            targetOU = Console.ReadLine().Trim();

            // Generate additional details
            firstInitial = Regex.Match(firstName, ".{1,1}").Value;
            lastInitial = Regex.Match(lastName, ".{1,1}").Value;
            username = $"{firstInitial.ToLower()}{lastName.ToLower()}";
            email = $"{username}@{myCompany}.com";
            password = $"New_User_{myCompany}_{firstInitial.ToUpper()}{lastInitial.ToUpper()}!";
            userProfile = $@"\\lmnas-02\users\{username}";
            clsUserFolder = $@"\\lmcls\sys\users\{firstInitial.ToLower()}{lastName.ToLower()}";

            Console.WriteLine($"\n-----------------------------------------------------------------------------------" +
                             $"\nFirst Name: {firstName}\n" +
                             $"Last Name: {lastName}\n" +
                             $"Display Name: {firstName} {lastName}\n" +
                             $"Username: {username}\n" +
                             $"Email Address: {email}\n" +
                             $"Temp Password: {password} \n" +
                             $"Department: {departmentEntry} \n" +
                             $"Title: {jobTitle} \n" +
                             $"Description: {description} \n" +
                             $"Physical Office: {office}" +
                             $"User Assgined OU: {targetOU} \n" +
                             $"Script Path: logon.bat \n" +
                             $"Home Drive: P: \n" +
                             $"User Home Directory: {userProfile} \n" +
                             $"CLS Folder Location: {clsUserFolder}\n" +
                             $"-----------------------------------------------------------------------------------\n");
           
            bool isExit = false;    
            while (!isExit)
            {
                Console.Write("Please verify all new user information are correct !!!(Y/N):");
                string confirmation = Console.ReadLine().ToUpper().Trim();
           
                if(confirmation == "Y")
                {
                    isExit = true;
                    Console.WriteLine("User information has been verified. \nCreating user...\n");
                }// end of if-statemnet
                else
                {
                    Console.WriteLine("\nReturning to menu....");
                    return;
                }// end of else-statement
            }// end of while
            try
            {
                string ouPath = $"LDAP://OU={targetOU},OU={myParentOU},DC={myDomain},DC={mydomainDotCom}";

                using (PrincipalContext context = new PrincipalContext(ContextType.Domain, null, adminUsername, adminPassword))
                {
                    
                    using (UserPrincipal user = new UserPrincipal(context))                                                                             // Creating new User
                    {
                        user.Name = $"{firstName} {lastName}";
                        user.SamAccountName = username;
                        user.UserPrincipalName = $"{username}@{myCompany}.com";
                        user.SetPassword(password);
                        user.GivenName = firstName;
                        user.Surname = lastName;
                        user.EmailAddress = email;
                        user.DisplayName = $"{firstName} {lastName}";
                        user.ScriptPath = "logon.bat";
                        user.Description = description;
                        user.HomeDrive = "P:";
                        user.HomeDirectory = userProfile;
                        user.Enabled = true;
                        user.UserCannotChangePassword = false;
                        user.PasswordNeverExpires = false;
                        user.Save();

                        using (DirectoryEntry userEntry = (DirectoryEntry)user.GetUnderlyingObject())                                                   // Move user to the specified OU
                        {
                            DirectoryEntry startOU = new DirectoryEntry(userEntry.Path);
                            DirectoryEntry endOU = new DirectoryEntry(ouPath, adminUsername, adminPassword);
                            userEntry.Properties["title"].Value = jobTitle;
                            userEntry.Properties["department"].Value = departmentEntry;
                            userEntry.Properties["physicalDeliveryOfficeName"].Value = office;
                            //userEntry.Properties["manager"].Value = manager;
                            userEntry.CommitChanges();
                            startOU.MoveTo(endOU);
                        }// end of using

                        Console.WriteLine($"User Account '{username}' Created Successfully!!!");
                        user.Dispose();
                    }// end of UserPrincipal using
                }// end of PrincipalContect using
                IsUserCreated(username);                                                                                                                // Verify account is created in AD
                AddNewUserToGroups(username, targetOU, adminUsername, adminPassword);                                                                   // Add using to basic groups based on select organizational unit (OU)
                CreateCLSFolder(clsUserFolder);                                                                                                         // Optional: Create CLS folder for new user
               CreateExchangeMailbox(adminUsername, adminPassword);                                                                                    // Create local Exchange mailbox
            }// end of try
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating user account: {ex.Message}");
            }// end of catch
        }// end of CreateUserAccount

        /// <summary>
        /// Check if a user account exists in Active Directory.
        /// </summary>
        /// <param name="username">The username to check.</param>
        /// <returns>True if the user exists, false otherwise.</returns>
        private bool IsUserCreated(string username)
        {
            Thread.Sleep(1000);
            try
            {
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
                    Console.WriteLine($"User Account Has Been Verfied: {user.DisplayName}!!!");
                    return user != null;
                }// end of using PrincipalContext 
            }// end of try
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking user account existence: {ex.Message}");
                return false;
            }// end of catach
        }// end of IsUserCreated
        /// <summary>
        /// Add the user to the appropriate groups based on the target OU.
        /// </summary>
        /// <param name="username">The username of the new user.</param>
        /// <param name="targetOu">The distinguished name of the target OU.</param>
        private void AddNewUserToGroups(string username, string targetOu, string adminUsername, string adminPassword)
        {
            Thread.Sleep(1000);
            string[] groups = null;
            string[] itGroups = { "_COLLECT", "_COLLECTKY", "_Training", "IT", "LM_IT" };
            string[] collectorGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Collectors", "LM_Collector", "NoOutboundEmail" };
            string[] adminStaffGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Administrative", "Staff" };
            string[] attyGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Attorneys", "LM_Atty" };
            string[] acctGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Accounting", "LM_Accounting", "NoAccountingEmail" };
            string[] complianceGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Compliance" };
            string[] michiganUsersGroups = { "_COLLECT", "CollectMI-11026982418", "_Training", "_Michigan", "MI_All_Users_Printers" };

            if (targetOu.Contains("IT")) groups = itGroups;
            else if (targetOu.Contains("Collector")) groups = collectorGroups;
            else if (targetOu.Contains("Admin Staff")) groups = adminStaffGroups;
            else if (targetOu.Contains("Atty")) groups = attyGroups;
            else if (targetOu.Contains("Acct")) groups = acctGroups;
            else if (targetOu.Contains("Compliance")) groups = complianceGroups;
            else if (targetOu.Contains("Michigan_Users")) groups = michiganUsersGroups;

            if (groups != null)
            {
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain, null, adminUsername, adminPassword))
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
                    if (user != null)
                    {
                        foreach (string groupName in groups)
                        {
                            GroupPrincipal group = GroupPrincipal.FindByIdentity(context, groupName);
                            if (group != null)
                            {
                                group.Members.Add(user);                                                                                                    // Adding user to groups based on selected OU
                                group.Save();
                            }// end of if statement
                        }// end of foreach
                        Console.WriteLine($"User '{username}' added to roups: {string.Join(", ", groups)}!!!");
                    }// end of if-statement
                    else
                    {
                        Console.WriteLine($"User '{username}' not found for group assignment.");
                    }// end of else-statement
                }// end of using PrincipalContext
            }// end of if-statement
            else
            {
                Console.WriteLine($"No group assignments found for the target OU '{targetOu}'");
            }// end of else-statement
        }// end of addUserToGroup
        private void CreateCLSFolder(string directoryPath)
        {
            Thread.Sleep(1000);
            if (!Directory.Exists(directoryPath))
            {
                try
                {
                    Directory.CreateDirectory(directoryPath);
                    Console.WriteLine($"CLS folder has been created in: {directoryPath}");
                }// end of try
                catch(Exception ex)
                {
                    Console.WriteLine($"An error has occured whie creating CLS folder: {ex.Message}");
                }// end of catch
            }// end of if-statement
            else
            {
                Console.WriteLine($"CLS file already Exist for this user: {username}");
            }// end of else-statement
        }// end of CreateCLSFolder

        private void LaunchBRPMgr()
        {

        }// end of LaunchBRPMgr

        // TODO - create mailbox
        private void CreateExchangeMailbox(string adminUsername, string adminPassword)
        {
            string createMailboxScript = $@"New-Mailbox -UserPrincipalName '{email}' -Alias '{username}' -Database '{myExhcangeDatabase}' -Name '{firstName} {lastName}' -Password ConvertTo-SecureString -String '{password}' -AsPlainText -Force";

            SecureString secureAdminPassword = new SecureString();
            foreach(char c in adminPassword.ToCharArray())
            {
                secureAdminPassword.AppendChar(c);
            }

            PSCredential adminCredential = new PSCredential(adminUsername, secureAdminPassword);


            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();

                using (PowerShell ps = PowerShell.Create())
                {
                    ps.Runspace = runspace;
                    ps.AddScript("Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force");
                    ps.Invoke();

                    
                    ps.AddScript($"${adminCredential} = New-Object System.Management.Automation.PSCredential ('{adminUsername}', (ConvertTo-SecureString -String '{adminPassword}' -AsPlainText -Force))");
                    ps.Invoke();
                
                    ps.AddScript("Import-Module ExchangeOnlineManagement");
                    ps.Invoke();
                    //  ps.AddScript($"{adminCredential} = New-Object System.Management.Automation.PSCredential ('{adminUsername}', ${secureAdminPassword})");
                    ps.AddScript(createMailboxScript);
                    Collection<PSObject> result = ps.Invoke();

                    try
                    {
                        //var result = ps.Invoke();
                        if (ps.HadErrors)
                        {
                            StringBuilder errorMessage = new StringBuilder("Error creating mailbox: ");
                            foreach (var error in ps.Streams.Error)
                            {
                                errorMessage.AppendLine(error.ToString());
                            }
                            Console.WriteLine(errorMessage.ToString());
                        }// end of if-statement
                        else
                        {
                            Console.WriteLine($"Mailbox created successfully for user '{username}'.");
                            
                        }// end of else-statement
                    }// end of try
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error creating mailbox: {ex.Message}");
                    }// end of catch
                }// end of using powershell
            }// end of using Runspace
        }// end of CreateExchangeMailbox
    }// end of class
}// end of namespace
