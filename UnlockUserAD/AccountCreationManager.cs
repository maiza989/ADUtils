using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Task = System.Threading.Tasks.Task;


// TODO - User Account Creation: Enable users to create new accounts in Active Directory. 
namespace UnlockUserAD
{
    public class AccountCreationManager
    {
        Program program;
        private PrincipalContext context;
        string firstName;
        string lastName;
        string jobTitle;
        string departmentEntry;
        string description;
       


        public AccountCreationManager() 
        {
            context = new PrincipalContext(ContextType.Domain);
            program = new Program();    
        }

        public void CreateUserAccount()
        {
            // Prompt for user details
            Console.Write("Enter new user's first name: ");
            firstName = Console.ReadLine();

            Console.Write("Enter new user's last name: ");
            string lastName = Console.ReadLine();

            Console.Write("Enter user's job title: ");
            string jobTitle = Console.ReadLine();

            Console.Write("Enter user's department: ");
            string departmentEntry = Console.ReadLine();

            Console.Write("Enter user description: ");
            string description = Console.ReadLine();

            Console.Write("Select New User OU (Admin Staff, Collector, Atty, Acct, IT, Compliance, Michigan_Users, Cooling_Users");
            string targetOU = Console.ReadLine().Trim();

            // Generate additional details
            string firstInitial = Regex.Match(firstName, ".{1,1}").Value.ToLower();
            string lastInitial = Regex.Match(lastName, ".{1,1}").Value.ToLower();
            string username = $"{firstInitial}{lastName.ToLower()}";
            string email = $"{username}@lloydmc.com";
            string password = $"New_User_lloydmc_{firstInitial}{lastInitial}!";

            

            string userProfile = $@"\\lmnas-02\users\{username}";
            string clsUserFolder = $@"\\lmcls\sys\users\{firstInitial}{lastName.ToLower()}";


            Console.WriteLine($"\n-----------------------------------------------------------------------------------" +
                             $"First Name: {firstName}\n" +
                             $"Last Name: {lastName}\n" +
                             $"Display Name: {firstName} {lastName}\n" +
                             $"Email Address: {email}\n" +
                             $"Temp Password: {password} \n" +
                             $"Department: {departmentEntry} \n" +
                             $"Description: {description} \n" +
                             $"User Assgined OU: {targetOU} \n" +
                             $"User Romaing Folder Location: {userProfile} \n" +
                             $"CLS Folder Location: {clsUserFolder}" +
                             $"-----------------------------------------------------------------------------------\n");
           
            bool isExit = false;    
            while (!isExit)
            {
                Console.Write("Please verify all new user information are correct !!!(Y/N)");
                string confirmation = Console.ReadLine().ToUpper().Trim();
           
                if(confirmation == "Y")
                {
                    isExit = true;
                    Console.WriteLine("User information has been verified. \nCreating user...");
                }// end of if-statemnet
                else
                {
                    Console.WriteLine("\nReturning to menu....");
                    return;
                }// end of else-statement
            }// end of while
            try
            {
                // Create a new UserPrincipal object
                using (UserPrincipal user = new UserPrincipal(context))
                {
                    DirectoryEntry directoryEntry = new DirectoryEntry($"LDAP://{targetOU}");
                  
                    DirectoryEntry userEntry = directoryEntry.Children.Add($"CN={firstName} {lastName}", "user");                                                   // Create a new DirectoryEntry object for the user

                    user.SamAccountName = username;
                    user.SetPassword(password);
                    user.GivenName = firstName;
                    user.Surname = lastName;
                    user.EmailAddress = email;
                    user.DisplayName = $"{firstName} {lastName}";
                    user.Description = description;
                    userEntry.Properties["title"].Value = jobTitle;
                    userEntry.Properties["department"].Value = departmentEntry;
                    user.HomeDirectory = userProfile;
                    user.Enabled = true;
                    user.UserCannotChangePassword = false;
                    user.PasswordNeverExpires = false;

                    // Save the new user to Active Directory
                    userEntry.CommitChanges();
                    userEntry.Invoke("SetPassword", new object[] {password});

                    user.Save();
                    user.Dispose();
                    Console.WriteLine($"User account '{username}' created successfully.");
                }// end of using
                IsUserCreated(username);
                AddNewUserToGroups(username, targetOU);
                CreateCLSFolder(clsUserFolder, username);
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
            Task.Delay(3000);
            try
            {
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
                    Console.WriteLine($"User Account is created: {user.DisplayName}");
                    return user != null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking user account existence: {ex.Message}");
                return false;
            }
        }


        /// <summary>
        /// Add the user to the appropriate groups based on the target OU.
        /// </summary>
        /// <param name="username">The username of the new user.</param>
        /// <param name="targetOu">The distinguished name of the target OU.</param>
        private void AddNewUserToGroups(string username, string targetOu)
        {
            Task.Delay(3000);
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
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
                    if (user != null)
                    {
                        foreach (string groupName in groups)
                        {
                            GroupPrincipal group = GroupPrincipal.FindByIdentity(context, groupName);
                            if (group != null)
                            {
                                group.Members.Add(user);
                                group.Save();
                            }
                        }
                        Console.WriteLine($"User '{username}' added to groups: {string.Join(", ", groups)}");
                        
                    }
                    else
                    {
                        Console.WriteLine($"User '{username}' not found for group assignment.");
                    }
                }
            }
            else
            {
                Console.WriteLine($"No group assignments found for the target OU '{targetOu}'");
            }
        }// end of addUserToGroup
        private void CreateCLSFolder(string directoryPath, string foldername)
        {
            string fullPath = Path.Combine(directoryPath, foldername);
            if (!Directory.Exists(fullPath))
            {
                try
                {
                    Directory.CreateDirectory(fullPath);
                    Console.WriteLine($"CLS folder has been created in: {fullPath}");
                    Task.Delay(3000);
                }// end of try
                catch(Exception ex)
                {
                    Console.WriteLine($"An error has occured whie creating CLS folder: {ex.Message}");
                }// end of catch
            }// end of if-statement
        }// end of CreateCLSFolder
        private void CreateLocalExchangeMailbox(string adminUsername, string adminPassword)
        {
            string exchangeUrl = "https://your_exchange_server/EWS/Exchange.asmx";
            string mailboxDatabase = "LMEX16DB1";

        // Authenticate with Exchange Server
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials(adminUsername, adminPassword);
            service.Url = new Uri(exchangeUrl);
            
            // Create Mailbox for the new user
            //service.EnableMailbox(username, mailboxDatabase);
        }

    }// end of class
}// end of namespace
