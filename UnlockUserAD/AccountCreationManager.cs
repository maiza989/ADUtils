using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Text.RegularExpressions;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Collections.ObjectModel;
using System.Diagnostics;
using Pastel;
using System.Drawing;
using Microsoft.Extensions.Configuration;
using System.Runtime.InteropServices;


// TODO - DONE User Account Creation: Enable users to create new accounts in Active Directory. 
// TODO - DONE Fix moving users to correct OU for MI and GA users. Need a switch case to select the correct OU parent. 
namespace ADUtils
{
    
    public class AccountCreationManager
    {

        EmailNotifcationManager emailNotification = new EmailNotifcationManager(Program.configuration);
        AuditLogManager auditLogManager;

        public readonly string _myDomain;
        public readonly string _myDomainDotCom;
        private string _myParentOU;
        public string _myCompany;
        private readonly string _myExchangeDatabase;
        private readonly string _myExchangeServer;

        

        public AccountCreationManager(AuditLogManager auditLogManager, IConfiguration configuration)
        {
            this.auditLogManager = auditLogManager;

            _myDomain = configuration["AccountCreationSettings:myDomain"];                                                      // Update with your domain
            _myDomainDotCom = configuration["AccountCreationSettings:myDomainDotCom"];                                          // Update with your second part of your domain (domain(.com))
            _myParentOU = configuration["AccountCreationSettings:myParentOU"];                                                  // Update with your path of Users OU
            _myCompany = configuration["AccountCreationSettings:myCompany"];                                                    // Update with your company email domain (*@companyName.com)
            _myExchangeDatabase = configuration["AccountCreationSettings:myExchangeDatabase"];                                  // Update with your Exchange server database
            _myExchangeServer = configuration["AccountCreationSettings:myExchangeServer"];                                      // Update with your exchange server name.

        }// end of constructor 

        public AccountCreationManager(IConfiguration configuration)
        {
            _myDomain = configuration["AccountCreationSettings:myDomain"];                                                      // Update with your domain
            _myDomainDotCom = configuration["AccountCreationSettings:myDomainDotCom"];                                          // Update with your second part of your domain (domain(.com))
            _myParentOU = configuration["AccountCreationSettings:myParentOU"];                                                  // Update with your path of Users OU
            _myCompany = configuration["AccountCreationSettings:myCompany"];                                                    // Update with your company email domain (*@companyName.com)
            _myExchangeDatabase = configuration["AccountCreationSettings:myExchangeDatabase"];                                  // Update with your Exchange server database
            _myExchangeServer = configuration["AccountCreationSettings:myExchangeServer"];                                      // Update with your exchange server name.
        }

        public AccountCreationManager() { }
        public int processSleepTimer = 1000;

        private string firstName;
        private string lastName;
        private string jobTitle;
        private string departmentEntry;
        private string description;
        private string office;
        private string manager;
        private string managerDN;
        private string targetOU;
        private string targetOUSelection;      
        private string firstInitial;
        private string lastInitial;
        private string username;
        private string email;
        private string password;
        private string userProfile;
        private string clsUserFolder;
        private string ouPath;
        

        private List<string> emailActionLog = new List<string>();                                                                // String list that hold email body
        /*
                /// <summary>
                /// Create a user in Active Directory based on the information provided by the user.
                /// </summary>
                /// <param name="adminUsername"></param>
                /// <param name="adminPassword"></param>
                public void CreateUserAccount(string adminUsername, string adminPassword)
                {
                    bool manualSteps = false;
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

                    Console.Write("Enter user office (KY, MI, GA, or Remote): ");
                    office = Console.ReadLine().Trim();

                    Console.Write("Enter user manager (SAM Account Name): ");
                    manager = Console.ReadLine();

                    //TODO - redo the selection of user creation This need to be do to create user with Horizon access if user is GA, MI or KY Remote
                    int choice;
                    bool validInput = false; 
                    switch (office.Trim().ToUpper())
                    {
                        case "KY":
                            do
                            {
                                Console.WriteLine($"Select New User OU. {"Enter Number".Pastel(Color.MediumPurple)}" +
                                                  $"\n 1- IT\n 2- Collector\n 3- Admin Staff\n 4- Atty\n 5- Acct\n 6- Compliance");
                                Console.Write("Enter your choice:");
                                validInput = int.TryParse(Console.ReadLine().Trim(), out choice) && choice >= 1 && choice <= 6;
                                if (!validInput) Console.WriteLine("Invalid input, please enter a number between 1 and 6.");
                            } while(!validInput);
                            targetOUSelection = choice.ToString();
                            break;
                        case "MI":
                            do
                            {
                                Console.WriteLine($"Select New User OU. {"Enter Number".Pastel(Color.MediumPurple)}" +
                                                  $"\n 7- Michigan Users\n 8- Michigan Collector\n 9- Michigan Admin staff \n 10- Michigan Attorney\n 11- Michigan Accounting");
                                Console.Write("Enter your choice:");
                                validInput = int.TryParse(Console.ReadLine().Trim(), out choice) && choice >= 7 && choice <=11;
                                if (!validInput) Console.WriteLine("Invalid input, please enter a number beteen 7 and 11.");
                            } while (!validInput);
                            targetOUSelection = choice.ToString();   
                            break;
                        case "GA":
                            do
                            {
                                Console.WriteLine($"Select New User OU. {"Enter Number".Pastel(Color.MediumPurple)}" +
                                                  $"\n 12- Default Georgia Users\n 13- Georgia Collector\n 14- Georgia Admin Staff\n 15- Georgia Atty\n 16- Georgia Accounting");
                                Console.Write("Enter your choice: ");
                                validInput = int.TryParse(Console.ReadLine().Trim(), out choice) && choice >= 12 && choice <= 16;
                                if (!validInput) Console.WriteLine("Invalid input, please enter a number between 12 and 16.");
                            } while (!validInput);
                            targetOUSelection = choice.ToString();
                            break;
                        case "REMOTE":
                            do
                            {
                                Console.WriteLine($"Select New User OU. {"Enter Number".Pastel(Color.MediumPurple)}" +
                                                  $"\n( 1- IT\n 2- Collector\n 3- Admin Staff\n 4- Atty\n 5- Acct\n 6- Compliance\n " +
                                                  $"7- Michigan Users\n 8- GA Users)\n");
                                Console.Write("Enter your choice: ");
                                validInput = int.TryParse(Console.ReadLine().Trim(), out choice) && choice >= 1 && choice <= 8;
                                if (!validInput) Console.WriteLine("Invalid input, please enter a number between 1 and 8.");
                            } while (!validInput);
                            targetOUSelection = choice.ToString();
                            break;
                        default:

                            Console.WriteLine($"Select New User OU. {"Enter Number".Pastel(Color.MediumPurple)}" +
                                              $"\n( 1- IT\n 2- Collector\n 3- Admin Staff\n 4- Atty\n 5- Acct\n 6- Compliance\n " +
                                              $"7- Michigan Users\n " +
                                              $"8- Georgia Users)\n");
                            Console.Write("Enter your choice:");
                            targetOUSelection = Console.ReadLine().Trim();
                            break;
                    }// end of switch

                    switch (targetOUSelection)
                    {
                                                                                // KY users
                        case "1":                                               // KY IT user   
                            targetOU = "IT";                
                            _myParentOU = "LloydMc_Lou";
                            break;
                        case "2":                                               // KY Collector user    
                            targetOU = "Collector";                             
                            _myParentOU = "LloydMc_Lou";
                            break;
                        case "3":                                               // KY Admin Staff user  
                            targetOU = "Admin Staff";       
                            _myParentOU = "LloydMc_Lou";
                            break;
                        case "4":                                               // KY Atty user   
                            targetOU = "Atty";
                            _myParentOU = "LloydMc_Lou";
                            break;
                        case "5":                                               // KY Accounting user       
                            targetOU = "Acct";
                            _myParentOU = "LloydMc_Lou";
                            break;
                        case "6":                                               // KY Compliance user   
                            targetOU = "Compliance";
                            _myParentOU = "LloydMc_Lou";
                            break;

                        case "7":                                               // General MI user
                            targetOU = "";
                            _myParentOU = "Michigan_Users";
                            break;
                        case "8":                                               // MI Collector User 
                            targetOU = "";
                            _myParentOU = "Michigan_Users";
                            break;
                        case "9":                                               // MI Admin Staff user  
                            targetOU = "";
                            _myParentOU = "Michigan_Users";
                            break;
                        case "10":                                              // MI Atty user 
                            targetOU = "";
                            _myParentOU = "Michigan_Users";
                            break;
                        case "11":                                              // MI Accounting user   
                            targetOU = "";
                            _myParentOU = "Michigan_Users";
                            break;


                        case "12":                                               // General GA user
                            targetOU = "";                                      // Empty for default location for GA users
                            _myParentOU = "Cooling_Users";
                            break;
                        case "13":                                               // GA Collector User
                            targetOU = "Call_Center";
                            _myParentOU = "Cooling_Users";
                            break;
                        case "14":                                              // GA Admin Staff user
                            targetOU = "GA_Staff";                                      // Empty since OU is not set up 
                            _myParentOU = "Cooling_Users";
                            break;
                        case "15":                                              // GA Atty user
                            targetOU = "";                                      // empty since OU is not set up
                            _myParentOU = "Cooling_Users";
                            break;
                        case "16":                                              // GA Accounting user
                            targetOU = "Accounting";
                            _myParentOU = "Cooling_Users";
                            break;
                        default:
                            targetOU = "Admin Staff";
                            break;
                    }// end of switch-case

                    // Generate additional details
                    firstInitial = Regex.Match(firstName, ".{1,1}").Value;
                    lastInitial = Regex.Match(lastName, ".{1,1}").Value;
                    username = $"{firstInitial.ToLower()}{lastName.ToLower()}";
                    email = $"{username}@{_myCompany}.com";
                    password = $"New_User_{_myCompany}_{firstInitial.ToUpper()}{lastInitial.ToUpper()}!";
                    userProfile = $@"\\lmnas-02\users\{username}";
                    clsUserFolder = $@"\\lmcls\sys\users\{firstInitial.ToLower()}{lastName.ToLower()}";

                    Console.WriteLine($"\n-----------------------------------------------------------------------------------" +
                                     $"\n{"First Name:",-20} {firstName}\n" +
                                     $"{"Last Name:",-20} {lastName}\n" +
                                     $"{"Display Name:",-20} {firstName} {lastName}\n" +
                                     $"{"Username:", -20} {username}\n" +
                                     $"{"Email Address:", -20} {email}\n" + 
                                     $"{"Temp Password:", -20} {password} \n" +
                                     $"{"Department:", -20} {departmentEntry} \n" +
                                     $"{"Title:", -20} {jobTitle} \n" +
                                     $"{"Description:", -20} {description} \n" +
                                     $"{"Physical Office:", -20} {office} \n" +
                                     $"{"User Assigned OU:", -20} {targetOU} \n" +
                                     $"{"User Parent OU", -20} {_myParentOU} \n" + 
                                     $"{"Script Path:", -20} logon.bat \n" +
                                     $"{"Home Drive:", -20} P: \n" +
                                     $"{"User Home Directory:", -20} {userProfile} \n" +
                                     $"{"CLS Folder Location:", -20} {clsUserFolder}\n" +
                                     $"-----------------------------------------------------------------------------------\n");

                    bool isExit = false;
                    while (!isExit)
                    {
                        Console.Write($"\nPlease verify all new user information are correct !!!{"(Y/N)".Pastel(Color.MediumPurple)}:");
                        string confirmation = Console.ReadLine().ToUpper().Trim();

                        if (confirmation == "Y")
                        {
                            isExit = true;
                            Console.WriteLine("User information has been verified. \nCreating user...\n");
                        }// end of if-statement
                        else
                        {
                            Console.WriteLine("\nReturning to menu....");
                            return;
                        }// end of else-statement
                    }// end of while

                    try
                    {
                        if (targetOU == "")
                        {
                           ouPath = $"LDAP://OU={_myParentOU},DC={_myDomain},DC={_myDomainDotCom}";

                        }// end of if statement
                        else
                        {
                           ouPath = $"LDAP://OU={targetOU},OU={_myParentOU},DC={_myDomain},DC={_myDomainDotCom}";
                        }// end of else statement
                        // TODO - DONE fix the traget and parent OU for GA and MI user. Currently the parent OU only works for KY users.

                        using (PrincipalContext context = new PrincipalContext(ContextType.Domain, null, adminUsername, adminPassword))
                        {  
                            using (UserPrincipal user = new UserPrincipal(context))                                                                     // Creating new User
                            {
                                user.Name = $"{firstName} {lastName}";
                                user.SamAccountName = username;
                                user.UserPrincipalName = $"{username}@{_myCompany}.com";
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

                                using (DirectoryEntry userEntry = (DirectoryEntry)user.GetUnderlyingObject())                                           // Move user to the specified OU
                                {
                                    DirectoryEntry startOU = new DirectoryEntry(userEntry.Path);
                                    DirectoryEntry endOU = new DirectoryEntry(ouPath, adminUsername, adminPassword);
                                    userEntry.Properties["title"].Value = jobTitle;
                                    userEntry.Properties["department"].Value = departmentEntry;
                                    userEntry.Properties["physicalDeliveryOfficeName"].Value = office;

                                    try
                                    {
                                        CheckManagerDN(); 
                                        userEntry.Properties["manager"].Value = managerDN;
                                    }
                                    catch(Exception ex)
                                    {
                                        Console.WriteLine($"An error occurred while filling manager field for the user: {ex.Message}");
                                    }
                                    userEntry.CommitChanges();
                                    try
                                    {
                                        startOU.MoveTo(endOU);

                                    }
                                    catch (COMException ex)
                                    {
                                        Console.WriteLine($"An error occurred while moving the user to the traget OU: {ex.Message}");
                                    }
                                }// end of using

                                Console.WriteLine($"User Account '{username}' Created Successfully!!!".Pastel(Color.DarkOliveGreen));

                                user.Dispose();


                            }// end of UserPrincipal using
                        }// end of PrincipalContect using

                        IsUserCreated(username);                                                                                                        // Verify account is created in AD
                        AddNewUserToGroups(username, region,role, adminUsername, adminPassword);                                                           // Add using to basic groups based on select organizational unit (OU)
                        CreateExchangeMailbox(adminUsername, adminPassword);                                                                            // Create local Exchange mailbox
                        CreateCLSFolder(clsUserFolder);                                                                                                 // Optional: Create CLS folder for new user
                        LaunchVLMMgr();                                                                                                                 // Optional: Open VLM to add CLS license to the user
                        LaunchPhoneSystemSite();                                                                                                        // Optional: Open RingCentral site to add EXT
                        LaunchO365Site();                                                                                                               // Optional: Open O365 site to add licensees to the user.

                        string logEntry = ($"New Account has been created \"{firstName} {lastName} | {username}\" in Active Directory\n " +
                                           $"\nUser added to {targetOU} OU and assgined basic groups \n" +
                                           $"\nNew Exchange MailBox has been created for \"{firstName} {lastName} | {username}\"\n" +
                                           $"\nNew CLS folder has been created for \"{firstName} {lastName} | {username}\"\n " +
                                           $"\nNew CLS license needs to be added manually for \"{firstName} {lastName} | {username}\"\n" +
                                           $"\nNew vMedia license needs to be added manually for \"{firstName} {lastName} | {username}\"\n" +
                                           $"\nNew EXT needs to be added manually for \"{firstName} {lastName} | {username}\"\n" +
                                           $"\nNew O365 license needs to be added manually for \"{firstName} {lastName} | {username}\"\n" +
                                           $"");
                        auditLogManager.Log(logEntry);
                        emailActionLog.Add(logEntry);
                        // TODO - Fix duplicate email log entries when creating multiple users in a row.    
                    }// end of try
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error creating user account: {ex.Message}");
                    }// end of catch

                    if (emailActionLog.Count > 0)
                    {
                        string emailBody = string.Join("\n", emailActionLog);
                        emailNotification.SendEmailNotification("ADUtil Action: Administrative Action in Active Directory", emailBody);
                    }// end of if statement
                    do
                    {
                        Console.Write($"Have you completed the manual steps for CLS, BRP, Phone, Office 365?{"(Y/N)".Pastel(Color.MediumPurple)}\nEnter your choice: ");
                        string result =  Console.ReadLine().Trim().ToUpper();
                        if(result == "Y")
                        {
                            Console.WriteLine("Manual Steps complete");
                            manualSteps = true;
                        }else if(result == "N")
                        {
                            Console.WriteLine("Manual Steps are reqired!!");
                        }
                    } while(!manualSteps);

                }// end of CreateUserAccount*/


        public void CreateUserAccount(string adminUsername, string adminPassword)
        {
            bool manualSteps = false;

            // -------------------------
            // Prompt for user details
            // -------------------------
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

            Console.Write("Enter user office (KY, MI, GA, or Remote): ");
            office = Console.ReadLine().Trim().ToUpper();

            Console.Write("Enter user manager (SAM Account Name): ");
            manager = Console.ReadLine();

            // -------------------------
            // Define OU options dynamically
            // -------------------------
            var TargetOUs = new List<TargetOU>
    {
        // KY
        new TargetOU("KY", "IT", "LloydMc_Lou"),
        new TargetOU("KY", "Collector", "LloydMc_Lou"),
        new TargetOU("KY", "Admin Staff", "LloydMc_Lou"),
        new TargetOU("KY", "Atty", "LloydMc_Lou"),
        new TargetOU("KY", "Acct", "LloydMc_Lou"),
        new TargetOU("KY", "Compliance", "LloydMc_Lou"),

        // MI
        new TargetOU("MI", "", "Michigan_Users", "Michigan Users"),
        new TargetOU("MI", "", "Michigan_Users", "Michigan Collector"),
        new TargetOU("MI", "", "Michigan_Users", "Michigan Admin Staff"),
        new TargetOU("MI", "", "Michigan_Users", "Michigan Atty"),
        new TargetOU("MI", "", "Michigan_Users", "Michigan Accounting"),

        // GA
        new TargetOU("GA", "", "Cooling_Users", "Default Georgia Users"),
        new TargetOU("GA", "Call_Center", "Cooling_Users", "Georgia Collector"),
        new TargetOU("GA", "GA_Staff", "Cooling_Users", "Georgia Admin Staff"),
        new TargetOU("GA", "", "Cooling_Users", "Georgia Atty"),
        new TargetOU("GA", "Accounting", "Cooling_Users", "Georgia Accounting"),

        // Remote
        new TargetOU("REMOTE", "IT", "LloydMc_Lou"),
        new TargetOU("REMOTE", "Collector", "LloydMc_Lou"),
        new TargetOU("REMOTE", "Admin Staff", "LloydMc_Lou"),
        new TargetOU("REMOTE", "Atty", "LloydMc_Lou"),
        new TargetOU("REMOTE", "Acct", "LloydMc_Lou"),
        new TargetOU("REMOTE", "Compliance", "LloydMc_Lou"),
        new TargetOU("REMOTE", "", "Michigan_Users", "Michigan Users"),
        new TargetOU("REMOTE", "", "Cooling_Users", "GA Users")
    };

            // Filter options by office
            var filteredOptions = TargetOUs.Where(o => o.Office == office).ToList();
            if (!filteredOptions.Any())
            {
                Console.WriteLine("Invalid office entered. Defaulting to KY IT.");
                filteredOptions = TargetOUs.Where(o => o.Office == "KY").Take(1).ToList();
            }

            // -------------------------
            // Display options and get selection
            // -------------------------
            Console.WriteLine($"Select New User OU. {"Enter Number".Pastel(Color.MediumPurple)}:");
            for (int i = 0; i < filteredOptions.Count; i++)
                Console.WriteLine($"{i + 1}- {filteredOptions[i].DisplayName}");

            int choice;
            bool validInput;
            do
            {
                Console.Write("Enter your choice: ");
                validInput = int.TryParse(Console.ReadLine().Trim(), out choice) && choice >= 1 && choice <= filteredOptions.Count;
                if (!validInput)
                    Console.WriteLine($"Invalid input, please enter a number between 1 and {filteredOptions.Count}.");
            } while (!validInput);

            var selectedOU = filteredOptions[choice - 1];
            targetOU = selectedOU.Role;
            _myParentOU = selectedOU.ParentOU;
            string region = selectedOU.Office; // for AddNewUserToGroups
            string role = selectedOU.Role;      // for AddNewUserToGroups

            // -------------------------
            // Generate username, email, etc.
            // -------------------------
            firstInitial = Regex.Match(firstName, ".{1,1}").Value;
            lastInitial = Regex.Match(lastName, ".{1,1}").Value;
            username = $"{firstInitial.ToLower()}{lastName.ToLower()}";
            email = $"{username}@{_myCompany}.com";
            password = $"New_User_{_myCompany}_{firstInitial.ToUpper()}{lastInitial.ToUpper()}!";
            userProfile = $@"\\lmnas-02\users\{username}";
            clsUserFolder = $@"\\lmcls\sys\users\{firstInitial.ToLower()}{lastName.ToLower()}";

            // -------------------------
            // Display summary
            // -------------------------
            Console.WriteLine($"\n-----------------------------------------------------------------------------------" +
                              $"\n{"First Name:",-20} {firstName}\n" +
                              $"{"Last Name:",-20} {lastName}\n" +
                              $"{"Display Name:",-20} {firstName} {lastName}\n" +
                              $"{"Username:",-20} {username}\n" +
                              $"{"Email Address:",-20} {email}\n" +
                              $"{"Temp Password:",-20} {password} \n" +
                              $"{"Department:",-20} {departmentEntry} \n" +
                              $"{"Title:",-20} {jobTitle} \n" +
                              $"{"Description:",-20} {description} \n" +
                              $"{"Physical Office:",-20} {office} \n" +
                              $"{"User Assigned OU:",-20} {targetOU} \n" +
                              $"{"User Parent OU",-20} {_myParentOU} \n" +
                              $"{"Script Path:",-20} logon.bat \n" +
                              $"{"Home Drive:",-20} P: \n" +
                              $"{"User Home Directory:",-20} {userProfile} \n" +
                              $"{"CLS Folder Location:",-20} {clsUserFolder}\n" +
                              $"-----------------------------------------------------------------------------------\n");

            // -------------------------
            // Confirm info
            // -------------------------
            while (true)
            {
                Console.Write($"\nPlease verify all new user information are correct !!!{"(Y/N)".Pastel(Color.MediumPurple)}: ");
                string confirmation = Console.ReadLine().ToUpper().Trim();
                if (confirmation == "Y")
                {
                    Console.WriteLine("User information has been verified. \nCreating user...\n");
                    break;
                }
                else
                {
                    Console.WriteLine("\nReturning to menu....");
                    return;
                }
            }

            // -------------------------
            // Create user in AD
            // -------------------------
            try
            {
                ouPath = string.IsNullOrEmpty(targetOU)
                    ? $"LDAP://OU={_myParentOU},DC={_myDomain},DC={_myDomainDotCom}"
                    : $"LDAP://OU={targetOU},OU={_myParentOU},DC={_myDomain},DC={_myDomainDotCom}";

                using (PrincipalContext context = new PrincipalContext(ContextType.Domain, null, adminUsername, adminPassword))
                {
                    using (UserPrincipal user = new UserPrincipal(context))
                    {
                        user.Name = $"{firstName} {lastName}";
                        user.SamAccountName = username;
                        user.UserPrincipalName = $"{username}@{_myCompany}.com";
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

                        using (DirectoryEntry userEntry = (DirectoryEntry)user.GetUnderlyingObject())
                        {
                            DirectoryEntry endOU = new DirectoryEntry(ouPath, adminUsername, adminPassword);
                            userEntry.Properties["title"].Value = jobTitle;
                            userEntry.Properties["department"].Value = departmentEntry;
                            userEntry.Properties["physicalDeliveryOfficeName"].Value = office;

                            try { CheckManagerDN(); userEntry.Properties["manager"].Value = managerDN; }
                            catch (Exception ex) { Console.WriteLine($"Manager error: {ex.Message}"); }

                            userEntry.CommitChanges();

                            try { new DirectoryEntry(userEntry.Path).MoveTo(endOU); }
                            catch (COMException ex) { Console.WriteLine($"Move error: {ex.Message}"); }
                        }


                        Console.WriteLine($"User Account '{username}' Created Successfully!!!".Pastel(Color.DarkOliveGreen));

                        user.Dispose();


                    }// end of UserPrincipal using
                }// end of PrincipalContect using

                IsUserCreated(username);                                                                                                        // Verify account is created in AD
                AddNewUserToGroups(username, region, role, adminUsername, adminPassword);                                                           // Add using to basic groups based on select organizational unit (OU)
                CreateExchangeMailbox(adminUsername, adminPassword);                                                                            // Create local Exchange mailbox
                CreateCLSFolder(clsUserFolder);                                                                                                 // Optional: Create CLS folder for new user
                LaunchVLMMgr();                                                                                                                 // Optional: Open VLM to add CLS license to the user
                LaunchPhoneSystemSite();                                                                                                        // Optional: Open RingCentral site to add EXT
                LaunchO365Site();                                                                                                               // Optional: Open O365 site to add licensees to the user.

                string logEntry = ($"New Account has been created \"{firstName} {lastName} | {username}\" in Active Directory\n " +
                                   $"\nUser added to {targetOU} OU and assgined basic groups \n" +
                                   $"\nNew Exchange MailBox has been created for \"{firstName} {lastName} | {username}\"\n" +
                                   $"\nNew CLS folder has been created for \"{firstName} {lastName} | {username}\"\n " +
                                   $"\nNew CLS license needs to be added manually for \"{firstName} {lastName} | {username}\"\n" +
                                   $"\nNew vMedia license needs to be added manually for \"{firstName} {lastName} | {username}\"\n" +
                                   $"\nNew EXT needs to be added manually for \"{firstName} {lastName} | {username}\"\n" +
                                   $"\nNew O365 license needs to be added manually for \"{firstName} {lastName} | {username}\"\n" +
                                   $"");
                auditLogManager.Log(logEntry);
                emailActionLog.Add(logEntry);
                // TODO - Fix duplicate email log entries when creating multiple users in a row.    
            }// end of try
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating user account: {ex.Message}");
            }// end of catch

            if (emailActionLog.Count > 0)
            {
                string emailBody = string.Join("\n", emailActionLog);
                emailNotification.SendEmailNotification("ADUtil Action: Administrative Action in Active Directory", emailBody);
            }// end of if statement
            do
            {
                Console.Write($"Have you completed the manual steps for CLS, BRP, Phone, Office 365?{"(Y/N)".Pastel(Color.MediumPurple)}\nEnter your choice: ");
                string result = Console.ReadLine().Trim().ToUpper();
                if (result == "Y")
                {
                    Console.WriteLine("Manual Steps complete");
                    manualSteps = true;
                }
                else if (result == "N")
                {
                    Console.WriteLine("Manual Steps are reqired!!");
                }
            } while (!manualSteps);

        }// end of CreateUserAccount*/

        /// <summary>
        /// Check if a user account exists in Active Directory.
        /// </summary>
        /// <param name="username">The username to check.</param>
        /// <returns>True if the user exists, false otherwise.</returns>
        private bool IsUserCreated(string username)
        {
            Thread.Sleep(processSleepTimer);
            try
            {
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);

                    Console.ForegroundColor = ConsoleColor.DarkCyan;
                    Console.WriteLine($"User Account Has Been Verfied: {user.DisplayName}!!!");
                    Console.ForegroundColor = ConsoleColor.Gray;

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
        /// Returns a string with a groups user is member of. 
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public string GetUserGroupsString(UserPrincipal user)
        {
            var userGroups = new List<string>();
            var groups = user.GetGroups();

            foreach (var group in groups)
            {
                userGroups.Add(group.Name);
            }

            string userGroupsString = string.Join(", ", userGroups);
            return userGroupsString;
        }// end of GetUserGroupsString



      /*  /// <summary>
        /// Add the user to the appropriate groups based on the target OU.
        /// </summary>
        /// <param name="username">The username of the new user.</param>
        /// <param name="targetOu">The distinguished name of the target OU.</param>
        private void AddNewUserToGroups(string username, string targetOu, string adminUsername, string adminPassword)
        {
            // TODO - Fix the group assignment for new hire to be dynamic
            // Section to add more group types
            Thread.Sleep(processSleepTimer);
            string[] groups = null;                                                                                                                                                                // Change the group var name and value to match your needs
            // KY
            string[] itGroups = { "_COLLECT", "_COLLECTKY", "_Training", "IT", "LM_IT" };                                                                                                           // 1
            string[] collectorGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Collectors", "LM_Collector", "NoOutboundEmail" };                                                                  // 2
            string[] adminStaffGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Administrative", "Staff" };                                                                                       // 3
            string[] attyGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Attorneys", "LM_Atty", "Duo_Users" };                                                                                   // 4
            string[] acctGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Accounting", "LM_Accounting", "NoAccountingEmail" };                                                                    // 5
            string[] complianceGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Compliance" };                                                                                                    // 6


            // MI
            string[] michiganUsersGroups = { "_COLLECT", "CollectMI-11026982418", "_Training", "_Michigan", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" };                                    // 7
            string[] michiganCollectorUsersGroups = { "_COLLECT", "CollectMI-11026982418", "_Training", "_Michigan", "Collectors", "LM_Collector",
                                                "NoOutboundEmail", "Horizon_Collector_RDS_Users", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" };                                              // 8                                   
            string[] michiganAdminStaffUsersGroups = { "_COLLECT", "CollectMI-11026982418", "_Training", "_Michigan", "Administrative",
                                                "Staff", "Horizon_RDS_Desktop_Users", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" };                                                          // 9                                    
            string[] michiganAttyUsersGroups = { "_COLLECT", "CollectMI-11026982418", "_Training", "_Michigan", "Attorneys",
                                                "LM_Atty", "Horizon_Attorney_RDS_Users", "MI_All_Users_Printers", "BRP_Staff_Horizon_User","Duo_Users","Deny_Outlook_OST_Redirection"  };           // 10                              
            string[] michiganAcctUsersGroups = { "_COLLECT", "CollectMI-11026982418", "_Training", "_Michigan", "Accounting","ACHCC_Full","Horizon_ACC_WycomeMI_Map",
                                                "Horizon_Accounting_RDS_Users","MI_Accounting_Printers", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" };                                       // 11

            // GA
            string[] georgiaUsersGroups = { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA" };                                                                                                 // 12
            string[] georgiaCollectorUsersGroups = { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA", "Collectors", "LM_Collector",
                                                "NoOutboundEmail", "Horizon_Collector_RDS_Users", "BRP_Staff_Horizon_User" };                                                                       // 13
            string[] georgiaAdminStaffUsersGroups = { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA", "Administrative",
                                                "Staff", "Horizon_RDS_Desktop_Users", "BRP_Staff_Horizon_User" };                                                                                   // 14
            string[] georgiaAttyUsersGroups = { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA", "Attorneys",
                                                "LM_Atty", "Georgia attorneys", "Horizon_Attorney_RDS_Users", "BRP_Staff_Horizon_User","Duo_Users","Deny_Outlook_OST_Redirection" };                // 15
            string[] georgiaAcctUsersGroups = { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA", "LM_Accounting",
                                                "NoAccountingEmail", "Horizon_Accounting_RDS_Users", "BRP_Staff_Horizon_User" };                                                                    // 16


            // Remote
            string[] KYRITGroups = { "_COLLECT", "_COLLECTKY", "_Training", "IT", "LM_IT", "Horizon_IT_User" };
            string[] KYRCollectorGorups = { "_COLLECT", "_COLLECTKY", "_Training", "Collectors", "LM_Collector", "NoOutboundEmail", "Horizon_Collector_RDS_Users" };
            string[] KYRadminStaffGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Administrative", "Staff", "Horizon_RDS_Desktop_Users" };
            string[] KYRAttyGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Attorneys", "LM_Atty", "Horizon_Attorney_RDS_Users", "Duo_Users", "Deny_Outlook_OST_Redirection" };
            string[] KYRAcctGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Accounting", "LM_Accounting", "NoAccountingEmail", "Horizon_Accounting_RDS_Users" };
            string[] KYRComplianceGroups = { "_COLLECT", "_COLLECTKY", "_Training", "Compliance", "Horizon_RDS_Desktop_Users" };

            // section to determine which group type the user is assigned. 
            // KY
            if (targetOu.Contains("IT") || targetOUSelection.Equals("1")) groups = itGroups;
            else if (targetOu.Contains("Collector") || targetOUSelection.Equals("2")) groups = collectorGroups;
            else if (targetOu.Contains("Admin Staff") || targetOUSelection.Equals("3")) groups = adminStaffGroups;
            else if (targetOu.Contains("Atty") || targetOUSelection.Equals("4")) groups = attyGroups;
            else if (targetOu.Contains("Acct") || targetOUSelection.Equals("5")) groups = acctGroups;
            else if (targetOu.Contains("Compliance") || targetOUSelection.Equals("6")) groups = complianceGroups;
            // MI
            else if (targetOu.Contains("Michigan_Users") || _myParentOU.Equals("Michigan_Users") && targetOUSelection.Equals("7")) groups = michiganUsersGroups;
            else if (targetOu.Contains("Michigan_Users") || _myParentOU.Equals("Michigan_Users") && targetOUSelection.Equals("8")) groups = michiganUsersGroups;
            else if (targetOu.Contains("Michigan_Users") || _myParentOU.Equals("Michigan_Users") && targetOUSelection.Equals("9")) groups = michiganUsersGroups;
            else if (targetOu.Contains("Michigan_Users") || _myParentOU.Equals("Michigan_Users") && targetOUSelection.Equals("10")) groups = michiganUsersGroups;
            else if (targetOu.Contains("Michigan_Users") || _myParentOU.Equals("Michigan_Users") && targetOUSelection.Equals("11")) groups = michiganUsersGroups;
            // GA
            else if (targetOu.Contains("Cooling_Users") || _myParentOU.Equals("Cooling_Users") && targetOUSelection.Equals("12")) groups = georgiaUsersGroups;
            else if (targetOu.Contains("Call_Center") || _myParentOU.Equals("Cooling_Users") && targetOUSelection.Equals("13")) groups = georgiaCollectorUsersGroups;
            else if (targetOu.Contains("GA_Staff") || _myParentOU.Equals("Cooling_Users") && targetOUSelection.Equals("14")) groups = georgiaAdminStaffUsersGroups;
            else if (targetOu.Contains("GA_Litigation") || _myParentOU.Equals("Cooling_Users") && targetOUSelection.Equals("15")) groups = georgiaAttyUsersGroups;
            else if (targetOu.Contains("Accounting") || _myParentOU.Equals("Cooling_Users") && targetOUSelection.Equals("16")) groups = georgiaAcctUsersGroups;

            //TODO - update the group assignment for new hire to be dynamic. 
            // List<string> groups = GroupAssignmentHelper.GetGroups(office, targetOu);

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
                        Console.WriteLine($"User '{username}' added to groups: {string.Join(", ", groups)}!!!".Pastel(Color.DarkOliveGreen));
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
        }// end of addUserToGroup*/


        /// <summary>
        /// Add the user to the appropriate groups based on the target OU.
        /// </summary>
        /// <param name="username">The username of the new user.</param>
        /// <param name="targetOu">The distinguished name of the target OU.</param>
        private void AddNewUserToGroups(string username, string region, string role, string adminUsername, string adminPassword)
        {
            // TODO - Fix the group assignment for new hire to be dynamic
            // Section to add more group types
            Thread.Sleep(processSleepTimer);
            //TODO - update the group assignment for new hire to be dynamic. 
            List<string> groups = GroupAssignmentHelper.GetGroups(region, role);
           
            if (groups != null && groups.Count > 0)
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
                                if (!group.Members.Contains(user))
                                {
                                group.Members.Add(user);                                                                                                    // Adding user to groups based on selected OU
                                group.Save();
                                }
                            }// end of if statement
                            else
                            {
                                Console.WriteLine($"Group '{groupName}' not found in Active Directory.");
                            }
                        }// end of foreach
                        Console.WriteLine($"User '{username}' added to groups: {string.Join(", ", groups)}!!!".Pastel(Color.DarkOliveGreen));
                    }// end of if-statement
                    else
                    {
                        Console.WriteLine($"User '{username}' not found for group assignment.");
                    }// end of else-statement
                }// end of using PrincipalContext
            }// end of if-statement
            else
            {
                Console.WriteLine($"No group assignments found for the Region='{region}', Role='{role}'");
            }// end of else-statement
        }// end of addUserToGroup



        /// <summary>
        /// Create a CLS folder in desired location
        /// </summary>
        /// <param name="directoryPath"></param>
        private void CreateCLSFolder(string directoryPath)
        {
            Thread.Sleep(processSleepTimer);
            if (!Directory.Exists(directoryPath))
            {
                try
                {
                    Directory.CreateDirectory(directoryPath);
                    Console.WriteLine($"CLS folder has been created in: {directoryPath}".Pastel(Color.DarkOliveGreen));
                }// end of try
                catch (Exception ex)
                {
                    Console.WriteLine($"An error has occured whie creating CLS folder: {ex.Message}");
                }// end of catch
            }// end of if-statement
            else
            {
                Console.WriteLine($"CLS file already Exist for this user: {username}".Pastel(Color.DarkGoldenrod));
            }// end of else-statement
        }// end of CreateCLSFolder

        /// <summary>
        /// Open BRP manager to create a BRP account for the new user manually.
        /// </summary>
        // TODO - DONE create mailbox
        /// <summary>
        /// A method that convert a string to a secure string.
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private SecureString StringToSecureString(string str)
        {
            SecureString secureString = new SecureString();
            foreach (char c in str.ToCharArray())
            {
                secureString.AppendChar(c);
            }// end of foreach
            secureString.MakeReadOnly();
            return secureString;
        }// End of StringToSecureString

        /// <summary>
        /// A method that prints our error from running powershell commands if any.
        /// </summary>
        /// <param name="ps"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        private bool HandlePowershellErrors(PowerShell ps, string action)
        {
            if (ps.Streams.Error.Count > 0)
            {
                foreach (ErrorRecord error in ps.Streams.Error)
                {
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine($"Error {action}: {error.Exception.Message}".Pastel(Color.IndianRed));
                    Console.ForegroundColor = ConsoleColor.Gray;
                }// end of foreach
                return true;
            }// end of if-statement
            return false;
        }// end of HandlePowershellErrors

        /// <summary>
        /// A method that runs mutliple powershell commands to create a mailbox the Exchange server.
        /// </summary>
        /// <param name="adminUsername"></param>
        /// <param name="adminPassword"></param>
        private void CreateExchangeMailbox(string adminUsername, string adminPassword)
        {
            Console.WriteLine("Creating User Mailbox...");
            Thread.Sleep(processSleepTimer);
            SecureString securePassword = StringToSecureString(adminPassword);

            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            using (PowerShell ps = PowerShell.Create())
            {
                ps.Runspace = runspace;

                // Set Execution Policy
                ps.AddCommand("Set-ExecutionPolicy");
                ps.AddParameter ("ExecutionPolicy", "RemoteSigned");

                ps.Invoke();
                if (HandlePowershellErrors(ps, "setting execution policy")) return;

                // Construct the New-PSSession command
                ps.Commands.Clear();
                ps.AddCommand("New-PSSession");
                ps.AddParameter("ConfigurationName", "Microsoft.Exchange");
                ps.AddParameter("ConnectionUri", new Uri($"http://{_myExchangeServer}/PowerShell/"));
                ps.AddParameter("Authentication", "Default");                                                   // use current context session to authenticate. 

               /* ps.AddParameter("Authentication", "Kerberos");
                ps.AddParameter("Credential", new PSCredential(adminUsername, securePassword));*/



                // Invoke New-PSSession to establish a session
                Collection<PSObject> result = ps.Invoke();
                if (HandlePowershellErrors(ps, "creating PSSession")) return;

                // Extract session
                var sessionId = result[0];

                // Import the session using Import-PSSession
                ps.Commands.Clear();
                ps.AddCommand("Import-PSSession");
                ps.AddParameter("Session", sessionId);
                ps.AddParameter("DisableNameChecking");

                ps.Invoke();
                if (HandlePowershellErrors(ps, "importing PSSession")) return;

                // Enable mailbox using Enable-Mailbox
                ps.Commands.Clear();
                ps.AddCommand("Enable-Mailbox");
                ps.AddParameter("Identity", username);
                ps.AddParameter("Database", _myExchangeDatabase);
                // TODO - Fix error about there is no avaliable globle catalog when enabling mailbox. 
               

                // Invoke Enable-Mailbox
                ps.Invoke();
                if (HandlePowershellErrors(ps, $"enabling mailbox for '{username}'")) return;

                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine($"Mailbox for '{username}' created successfully!!".Pastel(Color.DarkOliveGreen));
                Console.ForegroundColor = ConsoleColor.Gray;

                runspace.Close();
                runspace.Dispose();
            }// end of using
        }// end of CreateExhangeMailbox
        /// <summary>
        /// A method that launch BRP manager to create an new account for the user
        /// </summary>
        private void LaunchBRPMgr()
        {
            Console.WriteLine($"Please create account in BRPMgr for the new user MANUALLY!!!\nOpening BRP manager...");
            Thread.Sleep(processSleepTimer);
            ProcessStartInfo startInfo = new ProcessStartInfo();                                                                                    // Create a new process start info
            startInfo.FileName = @"F:\Imaging\BRPUserMgr.exe";                                                                                      // Set the file name to the path of the executable

            try
            {
                Process process = Process.Start(startInfo);                                                                                         // Start the process
            }// end of try
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while trying to start the process: {ex.Message}");
            }// end of catch
        }// end of LaunchBRPMgr
        /// <summary>
        /// A method the launch VLM to assigned CLS license for the users 
        /// </summary>
        private void LaunchVLMMgr()
        {
            Console.WriteLine($"Please add a CLS license to the new user MANUALLY!!!\nOpening VLM...");
            Thread.Sleep(processSleepTimer);

            ProcessStartInfo startInfo = new ProcessStartInfo();                                                                                    // Create a new process start info
            startInfo.FileName = @"F:\Vertican\VLM\licensemanager.exe";                                                                             // Set the file name to the path of the executable

            try
            {
                Process process = Process.Start(startInfo);                                                                                         // Start the process
            }// end of try
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while trying to start the process: {ex.Message}");
            }// end of catch
        }// end of LaunchVLMMgr
        /// <summary>
        /// A method that launch HostMyCalls site to set the phone for the user
        /// </summary>
        private void LaunchPhoneSystemSite()
        {
            Console.WriteLine($"Please Add a extension to the new user MANUALLY!!!\nOpening HostMyCalls Site...");
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "https://login.ringcentral.com/",
                    UseShellExecute = true // This is necessary to open the URL in the default browser
                };
                Process.Start(psi);
            }// end of try
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }// end of catch
        }// end of LaunchHostMyCallsSite
        /// <summary>
        /// A method that launch O365 site to set up licenses for the user if needed
        /// </summary>
        private void LaunchO365Site()
        {
            Console.WriteLine($"Please Add a O365 licnese to the new user MANUALLY!!!\nOpening Microsoft Office Site...");
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "https://www.office.com",
                    UseShellExecute = true // This is necessary to open the URL in the default browser
                };
                Process.Start(psi);
            }// end of try
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }// end of catch
        }// end of LaunchO365Site
        static void AnimateLine(string line)
        {
            foreach (char c in line)
            {
                Console.Write(c);
                Thread.Sleep(10); // Adjust delay for speed of animation
            }
            Console.WriteLine();
        }
        private void CheckManagerDN()
        {
            using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
            {
                using (var managerPrincipal = UserPrincipal.FindByIdentity(context, manager))
                {
                    if (managerPrincipal != null)
                    {
                        // Retrieve the underlying DirectoryEntry for the manager
                        DirectoryEntry directoryEntry = (DirectoryEntry)managerPrincipal.GetUnderlyingObject();

                        // Get the Distinguished Name (DN)
                        managerDN = directoryEntry.Properties["distinguishedName"].Value.ToString();
                        Console.WriteLine($"Manager found: {managerPrincipal.DisplayName} (DN: {managerDN})");
                    }
                }
            }
        }// end of CheckManagerDN
    }// end of class
}// end of namespace
