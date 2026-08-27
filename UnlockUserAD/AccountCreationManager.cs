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

        private readonly IConfiguration _configuration;

        public readonly string _myDomain;
        public readonly string _myDomainDotCom;

        /// <summary>Working parent OU — reassigned per selection while creating an account.</summary>
        private string _myParentOU;

        /// <summary>The configured parent OU, never reassigned. Use this for OU paths.</summary>
        public readonly string _myConfiguredParentOU;

        public string _myCompany;
        public readonly string _myExEmployeeOU;
        public readonly int _myDeletionGraceDays;


        public AccountCreationManager(AuditLogManager auditLogManager, IConfiguration configuration)
            : this(configuration)
        {
            this.auditLogManager = auditLogManager;
        }// end of constructor

        public AccountCreationManager(IConfiguration configuration)
        {
            _configuration = configuration;

            _myDomain = configuration["AccountCreationSettings:myDomain"];                                                      // Update with your domain
            _myDomainDotCom = configuration["AccountCreationSettings:myDomainDotCom"];                                          // Update with your second part of your domain (domain(.com))
            _myParentOU = configuration["AccountCreationSettings:myParentOU"];                                                  // Update with your path of Users OU
            _myConfiguredParentOU = _myParentOU;
            _myCompany = configuration["AccountCreationSettings:myCompany"];                                                    // Update with your company email domain (*@companyName.com)
            _myExEmployeeOU = configuration["AccountCreationSettings:myExEmployeeOU"] ?? "Ex Employee";                         // OU that deactivated accounts are moved into
            _myDeletionGraceDays = int.TryParse(configuration["AccountCreationSettings:myDeletionGraceDays"], out int days) ? days : 31;
        }
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
       // private string targetOUSelection;      
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
                            Console.WriteLine("User information has been verified. \nCreating user...\n".Pastel(Color.DarkCyan));
                        }// end of if-statement
                        else
                        {
                            Console.WriteLine("\nReturning to menu....".Pastel(Color.Gray));
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


        /// <summary>
        /// Prompts until a non-empty name is entered, or returns null if the operator types 'exit'.
        /// Empty names previously produced an empty sAMAccountName and a confusing Save() failure.
        /// </summary>
        private string PromptForRequiredName(string prompt, string fieldLabel)
        {
            while (true)
            {
                AppLog.Prompt(prompt);
                string value = ConsoleInput.ReadTrimmed();

                if (value.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    AppLog.Screen("\nReturning to menu....", Color.Gray);
                    return null;
                }
                if (value.Length == 0)
                {
                    AppLog.Warn($"{fieldLabel} is required (or type {"'exit'".Pastel(Color.MediumPurple)} to return to the menu).", color: Color.IndianRed);
                    continue;
                }
                // These land in the CN and the DN. Reject the RFC 4514 special characters rather
                // than emitting a malformed DN that fails with an opaque COM error.
                int bad = value.IndexOfAny(new[] { ',', '+', '"', '\\', '<', '>', ';', '=', '#', '/' });
                if (bad >= 0)
                {
                    AppLog.Warn($"{fieldLabel} cannot contain '{value[bad]}' — it is reserved in Active Directory names.", color: Color.IndianRed);
                    continue;
                }
                return value;
            }
        }// end of PromptForRequiredName

        /// <summary>
        /// Builds a random temporary password.
        ///
        /// This replaced a derived pattern ("New_User_{company}_{INITIALS}!") that anyone who knew
        /// the convention could guess from the new hire's name. Guaranteed to satisfy the same
        /// rules <see cref="PasswordManager.IsPasswordVaild"/> enforces: length, upper, lower,
        /// digit and symbol.
        /// </summary>
        private static string GenerateTempPassword()
        {
            const string upper = "ABCDEFGHJKLMNPQRSTUVWXYZ";       // no I/O — avoids read-back confusion
            const string lower = "abcdefghijkmnopqrstuvwxyz";      // no l
            const string digits = "23456789";                      // no 0/1
            const string symbols = "!@#$%^&*?-_";
            const string all = upper + lower + digits + symbols;
            int length = PasswordManager.MinimumPasswordLength + 1;  // one over the enforced minimum

            var chars = new List<char>(length)
            {
                // One from each class up front so the result always satisfies the policy.
                PickRandom(upper),
                PickRandom(lower),
                PickRandom(digits),
                PickRandom(symbols)
            };
            while (chars.Count < length)
            {
                chars.Add(PickRandom(all));
            }

            // Fisher-Yates with a cryptographic RNG, so the per-class characters aren't
            // always in positions 0-3.
            for (int i = chars.Count - 1; i > 0; i--)
            {
                int j = System.Security.Cryptography.RandomNumberGenerator.GetInt32(i + 1);
                (chars[i], chars[j]) = (chars[j], chars[i]);
            }
            return new string(chars.ToArray());
        }// end of GenerateTempPassword

        private static char PickRandom(string set)
        {
            return set[System.Security.Cryptography.RandomNumberGenerator.GetInt32(set.Length)];
        }// end of PickRandom

        public void CreateUserAccount(string adminUsername, string adminPassword)
        {
            bool manualSteps = false;

            emailActionLog.Clear(); // Add at top of CreateUserAccount
            // -------------------------
            // Prompt for user details
            // -------------------------
            firstName = PromptForRequiredName("Enter new user's first name: ", "First name");
            if (firstName == null) return;

            lastName = PromptForRequiredName("Enter new user's last name: ", "Last name");
            if (lastName == null) return;

            AppLog.Prompt("Enter user's job title: ");
            jobTitle = ConsoleInput.ReadTrimmed();

            AppLog.Prompt("Enter user's department: ");
            departmentEntry = ConsoleInput.ReadTrimmed();

            AppLog.Prompt("Enter user description: ");
            description = ConsoleInput.ReadTrimmed();

            AppLog.Prompt("Enter user office (KY, MI, GA, or Remote): ");
            office = ConsoleInput.ReadTrimmedUpper();

            AppLog.Prompt("Enter user manager (SAM Account Name): ");
            manager = ConsoleInput.ReadTrimmed();

            // -------------------------
            // Define OU options dynamically
            // -------------------------
            // Shared with the "Validate Group Assignment" report, so the rows the operator picks
            // from are exactly the rows that get checked against AD.
            var TargetOUs = GroupAssignmentHelper.GetTargetOUs();

            // Filter options by office
            var filteredOptions = TargetOUs.Where(o => o.Office == office).ToList();
            if (!filteredOptions.Any())
            {
                AppLog.Warn($"'{office}' is not a recognized office. Valid values: " +
                            $"{string.Join(", ", TargetOUs.Select(o => o.Office).Distinct())}.", color: Color.IndianRed);
                AppLog.Screen("Returning to menu — no account was created.", Color.Gray);
                return;
            }

            // -------------------------
            // Display options and get selection
            // -------------------------
            AppLog.Screen($"Select New User OU. {"Enter Number".Pastel(Color.MediumPurple)}:");
            for (int i = 0; i < filteredOptions.Count; i++)
                AppLog.Screen($"{i + 1}- {filteredOptions[i].DisplayName}");

            int choice;
            bool validInput;
            do
            {
                AppLog.Prompt("Enter your choice: ");
                validInput = int.TryParse(ConsoleInput.ReadTrimmed(), out choice) && choice >= 1 && choice <= filteredOptions.Count;
                if (!validInput)
                    AppLog.Warn($"Invalid input, please enter a number between 1 and {filteredOptions.Count}.", color: Color.IndianRed);
            } while (!validInput);

            var selectedOU = filteredOptions[choice - 1];
            // SubOu, not Role: the AD OU is often named differently from the role label
            // ("Atty" -> OU=Attorneys, "GA_Staff" -> OU=Staff), and empty means "sit in the parent".
            targetOU = selectedOU.SubOu;
            _myParentOU = selectedOU.ParentOU;
            string region = selectedOU.Region;   // not Office — see TargetOU.Region
            string role = selectedOU.Role; // keep "General" for group lookup

            // Fail before creating anything if this region/role has no group mapping, rather
            // than creating an account that silently lands with no group membership.
            var plannedGroups = GroupAssignmentHelper.GetGroups(region, role);
            if (plannedGroups.Count == 0)
            {
                AppLog.Warn($"No group assignment is defined for Region='{region}', Role='{role}'.", color: Color.IndianRed);
                AppLog.Screen("Add it to GroupAssignmentHelper before creating this account. Nothing was created.", Color.Gray);
                return;
            }

            // -------------------------
            // Generate username, email, etc.
            // -------------------------
            firstInitial = firstName.Substring(0, 1);
            lastInitial = lastName.Substring(0, 1);
            username = $"{firstInitial.ToLower()}{lastName.ToLower()}";
            email = $"{username}@{_myCompany}.com";
            password = $"New_User_lloydmc_{firstInitial}{lastInitial}!";
            userProfile = $@"\\lmusrdata\User_Profiles\{username}";
            clsUserFolder = $@"\\lmcls\sys\users\{firstInitial.ToLower()}{lastName.ToLower()}";

            // sAMAccountName is limited to 20 characters in AD; Save() would fail with an
            // unhelpful COM error otherwise.
            if (username.Length > 20)
            {
                AppLog.Error($"Generated username '{username}' is {username.Length} characters; " +
                             "AD allows a maximum of 20. Shorten the last name or create this account manually.");
                return;
            }

            // -------------------------
            // Display summary
            // -------------------------
            // Split deliberately: the temp password goes to the screen ONLY. Everything else is
            // logged. Do not merge these back into one call -- it would write the generated
            // password into the log files.
            AppLog.Screen($"\n-----------------------------------------------------------------------------------" +
                          $"\n{"First Name:",-20} {firstName}\n" +
                          $"{"Last Name:",-20} {lastName}\n" +
                          $"{"Display Name:",-20} {firstName} {lastName}\n" +
                          $"{"Username:",-20} {username}\n" +
                          $"{"Email Address:",-20} {email}");

            AppLog.ScreenOnly($"{"Temp Password:",-20} {password}");                                 // never logged

            AppLog.Screen($"{"Department:",-20} {departmentEntry} \n" +
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
                AppLog.Prompt($"\nPlease verify all new user information are correct !!!{"(Y/N)".Pastel(Color.MediumPurple)}: ");
                string confirmation = ConsoleInput.ReadTrimmedUpper();
                if (confirmation == "Y")
                {
                    AppLog.Info("User information has been verified. \nCreating user...\n", Color.DarkCyan);
                    break;
                }
                else
                {
                    AppLog.Screen("\nReturning to menu....", Color.Gray);
                    return;
                }
            }

            // -------------------------
            // Create user in AD
            // -------------------------
            bool accountCreated = false;
            bool movedToTargetOU = false;
            bool managerSet = false;
            bool mailboxCreated = false;
            try
            {
                // Built by the same helper the validation report uses, so a clean report means the
                // creator really does target the OU that was checked.
                ouPath = $"LDAP://{GroupAssignmentHelper.BuildRelativeOuPath(selectedOU)},DC={_myDomain},DC={_myDomainDotCom}";

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

                        // The temp password is handed over verbally, so require a change at first
                        // logon rather than leaving it valid for the full domain max password age.
                        try
                        {
                            user.ExpirePasswordNow();
                        }
                        catch (Exception ex)
                        {
                            AppLog.Warn($"Could not flag the password for change at next logon: {ex.Message}", ex, Color.DarkGoldenrod);
                            AppLog.Warn("Set 'User must change password at next logon' manually in ADUC.", color: Color.DarkGoldenrod);
                        }

                        // GetUnderlyingObject() returns an object owned by the UserPrincipal --
                        // don't dispose it here, or user.Dispose() operates on a disposed entry.
                        DirectoryEntry userEntry = (DirectoryEntry)user.GetUnderlyingObject();
                        using (DirectoryEntry endOU = new DirectoryEntry(ouPath, adminUsername, adminPassword))
                        {
                            userEntry.Properties["title"].Value = jobTitle;
                            userEntry.Properties["department"].Value = departmentEntry;
                            userEntry.Properties["physicalDeliveryOfficeName"].Value = office;

                            try
                            {
                                CheckManagerDN();
                                userEntry.Properties["manager"].Value = managerDN;
                                managerSet = true;
                            }
                            catch (Exception ex)
                            {
                                AppLog.Warn($"Manager could not be set to '{manager}': {ex.Message}", ex, Color.DarkGoldenrod);
                            }

                            userEntry.CommitChanges();

                            try
                            {
                                using (var startOU = new DirectoryEntry(userEntry.Path))
                                {
                                    startOU.MoveTo(endOU);
                                }
                                movedToTargetOU = true;
                            }
                            catch (COMException ex)
                            {
                                AppLog.Error($"Move to '{ouPath}' FAILED: {ex.Message}", ex, Color.Crimson);
                                AppLog.Warn($"'{username}' exists but is still in the default Users container — move it manually.", color: Color.Crimson);
                            }
                        }

                        AppLog.Info($"User Account '{username}' Created Successfully!!!", Color.DarkOliveGreen);
                    }// end of UserPrincipal using
                }// end of PrincipalContect using

                IsUserCreated(username);                                                                                                        // Verify account is created in AD
                AddNewUserToGroups(username, region, role, adminUsername, adminPassword);                                                           // Add using to basic groups based on select organizational unit (OU)
                mailboxCreated = CreateExchangeMailbox(adminUsername, adminPassword);                                                           // Create local Exchange mailbox
                CreateCLSFolder(clsUserFolder);                                                                                                 // Optional: Create CLS folder for new user
                LaunchVLMMgr();                                                                                                                 // Optional: Open VLM to add CLS license to the user
                LaunchPhoneSystemSite();                                                                                                        // Optional: Open RingCentral site to add EXT
                LaunchO365Site();                                                                                                               // Optional: Open O365 site to add licensees to the user.

                // Report only what actually happened. Previously this claimed the OU move and the
                // mailbox succeeded even when both had failed and been swallowed above.
                string who = $"\"{firstName} {lastName} | {username}\"";
                var log = new List<string>
                {
                    $"New Account has been created {who} in Active Directory",
                    movedToTargetOU
                        ? $"User moved to {(string.IsNullOrEmpty(targetOU) ? _myParentOU : targetOU)} OU and assigned basic groups"
                        : $"*** MOVE FAILED *** {who} is still in the default Users container and must be moved manually. Basic groups were assigned.",
                    mailboxCreated
                        ? $"New Exchange MailBox has been created for {who}"
                        : $"*** MAILBOX NOT CREATED *** Exchange mailbox for {who} must be created manually.",
                    $"New CLS folder has been created for {who}",
                    $"New CLS license needs to be added manually for {who}",
                    $"New vMedia license needs to be added manually for {who}",
                    $"New EXT needs to be added manually for {who}",
                    $"New O365 license needs to be added manually for {who}"
                };
                if (!managerSet)
                {
                    log.Add($"*** MANAGER NOT SET *** manager '{manager}' could not be applied to {who}.");
                }

                string logEntry = string.Join("\n\n", log) + "\n";
                auditLogManager.Log(logEntry);
                emailActionLog.Add(logEntry);
                accountCreated = true; // ← mark creation as successful
            }// end of try
            catch (Exception ex)
            {
                AppLog.Error($"Error creating user account: {ex.Message}", ex, Color.IndianRed);
            }// end of catch

            if (!accountCreated) return; // only exits if creation actually failed

            if (emailActionLog.Count > 0)
            {
                string emailBody = string.Join("\n", emailActionLog);
                emailNotification.SendEmailNotification("ADUtil Action: Administrative Action in Active Directory", emailBody);
            }// end of if statement
            do
            {
                AppLog.Prompt($"Have you completed the manual steps for CLS, BRP, Phone, Office 365?{"(Y/N)".Pastel(Color.MediumPurple)}\nEnter your choice: ");
                string result = ConsoleInput.ReadTrimmedUpper();
                if (result == "Y")
                {
                    AppLog.Info("Manual Steps complete", Color.DarkOliveGreen);
                    manualSteps = true;
                }
                else if (result == "N")
                {
                    AppLog.Warn("Manual Steps are reqired!!", color: Color.DarkGoldenrod);
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
                using (PrincipalContext context = AdminSession.CreateContext())
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
                    if(user != null)
                    {
                        AppLog.Info($"User Account Has Been Verified: {user.DisplayName}!!!", Color.DarkCyan);
                        return true;
                    }
                    return false;
                }// end of using PrincipalContext 
            }// end of try
            catch (Exception ex)
            {
                AppLog.Error($"Error checking user account existence: {ex.Message}", ex, Color.IndianRed);
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
                                AppLog.Warn($"Group '{groupName}' not found in Active Directory.", color: Color.DarkGoldenrod);
                            }
                        }// end of foreach
                        AppLog.Info($"User '{username}' added to groups: {string.Join(", ", groups)}!!!", Color.DarkOliveGreen);
                    }// end of if-statement
                    else
                    {
                        AppLog.Warn($"User '{username}' not found for group assignment.", color: Color.IndianRed);
                    }// end of else-statement
                }// end of using PrincipalContext
            }// end of if-statement
            else
            {
                AppLog.Warn($"No group assignments found for the Region='{region}', Role='{role}'", color: Color.IndianRed);
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
                    AppLog.Info($"CLS folder has been created in: {directoryPath}", Color.DarkOliveGreen);
                }// end of try
                catch (Exception ex)
                {
                    AppLog.Error($"An error has occured whie creating CLS folder: {ex.Message}", ex);
                }// end of catch
            }// end of if-statement
            else
            {
                AppLog.Warn($"CLS file already Exist for this user: {username}", color: Color.DarkGoldenrod);
            }// end of else-statement
        }// end of CreateCLSFolder

        /// <summary>
        /// Open BRP manager to create a BRP account for the new user manually.
        /// </summary>

        /// <summary>
        /// Enables an on-prem Exchange mailbox for the newly created user.
        /// </summary>
        /// <returns>
        /// True only when Enable-Mailbox completed without errors. Previously this always printed
        /// "created successfully" even when the cmdlet had failed, and the audit log and
        /// notification email inherited that false claim.
        /// </returns>
        private bool CreateExchangeMailbox(string adminUsername, string adminPassword)
        {
            AppLog.Info("Creating User Mailbox...", Color.DarkCyan);
            Thread.Sleep(processSleepTimer);

            using (var exchange = new ExchangeSessionManager(_configuration))
            {
                if (!exchange.Connect())
                {
                    AppLog.Warn($"Mailbox for '{username}' was NOT created — create it manually in Exchange.", color: Color.Crimson);
                    return false;
                }

                var parameters = new Dictionary<string, object>
                {
                    ["Identity"] = username,
                    ["Database"] = exchange.ExchangeDatabase
                };
                // Pin to a known DC. Letting Exchange choose is what produced the
                // "no available global catalog" failure.
                if (exchange.DomainController != null)
                {
                    parameters["DomainController"] = exchange.DomainController;
                }

                if (!exchange.RunCommand("Enable-Mailbox", $"enabling mailbox for '{username}'", parameters))
                {
                    AppLog.Warn($"Mailbox for '{username}' was NOT created — create it manually in Exchange.", color: Color.Crimson);
                    return false;
                }

                AppLog.Info($"Mailbox for '{username}' created successfully!!", Color.DarkOliveGreen);
                return true;
            }// end of using
        }// end of CreateExhangeMailbox
        /// <summary>
        /// A method that launch BRP manager to create an new account for the user
        /// </summary>
        private void LaunchBRPMgr()
        {
            AppLog.Screen($"Please create account in BRPMgr for the new user MANUALLY!!!\nOpening BRP manager...");
            Thread.Sleep(processSleepTimer);
            ProcessStartInfo startInfo = new ProcessStartInfo();                                                                                    // Create a new process start info
            startInfo.FileName = @"F:\Imaging\BRPUserMgr.exe";                                                                                      // Set the file name to the path of the executable

            try
            {
                Process process = Process.Start(startInfo);                                                                                         // Start the process
            }// end of try
            catch (Exception ex)
            {
                AppLog.Error($"An error occurred while trying to start the process: {ex.Message}", ex);
            }// end of catch
        }// end of LaunchBRPMgr
        /// <summary>
        /// A method the launch VLM to assigned CLS license for the users 
        /// </summary>
        private void LaunchVLMMgr()
        {
            AppLog.Screen($"Please add a CLS license to the new user MANUALLY!!!\nOpening VLM...");
            Thread.Sleep(processSleepTimer);

            ProcessStartInfo startInfo = new ProcessStartInfo();                                                                                    // Create a new process start info
            startInfo.FileName = @"F:\Vertican\VLM\licensemanager.exe";                                                                             // Set the file name to the path of the executable

            try
            {
                Process process = Process.Start(startInfo);                                                                                         // Start the process
            }// end of try
            catch (Exception ex)
            {
                AppLog.Error($"An error occurred while trying to start the process: {ex.Message}", ex);
            }// end of catch
        }// end of LaunchVLMMgr
        /// <summary>
        /// A method that launch HostMyCalls site to set the phone for the user
        /// </summary>
        private void LaunchPhoneSystemSite()
        {
            AppLog.Screen($"Please Add a extension to the new user MANUALLY!!!\nOpening HostMyCalls Site...");
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
                AppLog.Error($"An error occurred: {ex.Message}", ex);
            }// end of catch
        }// end of LaunchHostMyCallsSite
        /// <summary>
        /// A method that launch O365 site to set up licenses for the user if needed
        /// </summary>
        private void LaunchO365Site()
        {
            AppLog.Screen($"Please Add a O365 licnese to the new user MANUALLY!!!\nOpening Microsoft Office Site...");
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
                AppLog.Error($"An error occurred: {ex.Message}", ex);
            }// end of catch
        }// end of LaunchO365Site
        static void AnimateLine(string line)
        {
            foreach (char c in line)
            {
                Console.Write(c);                                                                    // character-by-character animation; logging each char would be noise
                Thread.Sleep(10); // Adjust delay for speed of animation
            }
            AppLog.Blank();
        }
        private void CheckManagerDN()
        {
            if (string.IsNullOrWhiteSpace(manager))
                throw new InvalidOperationException("Manager SAM account name is empty.");

            using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
            using (var managerPrincipal = UserPrincipal.FindByIdentity(context, manager))
            {
                if (managerPrincipal == null)
                    throw new InvalidOperationException($"Manager '{manager}' was not found in Active Directory.");

                DirectoryEntry directoryEntry = (DirectoryEntry)managerPrincipal.GetUnderlyingObject();
                managerDN = directoryEntry.Properties["distinguishedName"].Value?.ToString();

                if (string.IsNullOrEmpty(managerDN))
                    throw new InvalidOperationException($"Could not retrieve DN for manager '{manager}'.");

                AppLog.Info($"Manager found: {managerPrincipal.DisplayName} (DN: {managerDN})", Color.DarkCyan);
            }
        }// end of CheckManagerDN
    }// end of class
}// end of namespace
