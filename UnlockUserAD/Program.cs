using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using ADUtils;
using Pastel;
using System.Drawing;
using Microsoft.Extensions.Configuration;
// TODO - Create/Delete Groups: Allow creating and deleting security groups or distribution lists.


/// <summary>
/// All deployment-specific values live in Appsettings.json — see Appsettings.example.json
/// for the full set of keys. Copy that file to Appsettings.json and fill it in before
/// running. Nothing in this project needs to be edited to point it at a domain.
/// </summary>
class Program
{

    public static string adminUsername;
    private static string adminPassword;
    static private bool isAuthenticated = false;
    public static IConfiguration configuration;


    static void GetAdminCreditials()
    {

        AppLog.Prompt("Enter admin username: ");
        adminUsername = ConsoleInput.ReadTrimmed();
        AppLog.Prompt("Enter admin password: ");
        adminPassword = PasswordManager.GetPassword().Trim();
    }


    static void Main(string[] args)
    {
        // Before any output: puts the console on UTF-8 so box-drawing and status glyphs render.
        // A legacy code page best-fit-maps them instead, which silently turned em dashes into
        // hyphens and would make the framing unreadable.
        ConsoleUi.Initialize();

        try
        {
            configuration = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("Appsettings.json", optional: false, reloadOnChange: true)
                .Build();
        }
        catch (FileNotFoundException)
        {
            AppLog.Error("Appsettings.json was not found.");
            AppLog.Screen($"Copy {"Appsettings.example.json".Pastel(Color.MediumPurple)} to " +
                          $"{"Appsettings.json".Pastel(Color.MediumPurple)} in {AppContext.BaseDirectory} " +
                          "and fill in your domain, Exchange and SMTP values.");
            return;
        }
        catch (Exception ex)
        {
            AppLog.Error($"Appsettings.json could not be read: {ex.Message}", ex, Color.IndianRed);
            return;
        }

        ActiveDirectoryManager ADManager = new ActiveDirectoryManager();
        AccountCreationManager ACManager;
        PasswordManager PWDManager = null;
        ADGroupActionManager ADGroupManager = null;
        AuditLogManager auditLogManager = null;
        AccountDeactivationManager ACCDeactivationManager = new AccountDeactivationManager();

        string _myDomainName = configuration["AccountCreationSettings:myDomainName"];

        AppLog.Screen("Starting Active Directory Manager...");
        if (AuditLogManager.VerifyLoggingConfigured())
        {
            AppLog.Screen($"Logs: {AuditLogManager.LogDirectory}", Color.DarkGray);
        }

        try
        {
        do
        {
            try
            {
                GetAdminCreditials();

                using (PrincipalContext context = new PrincipalContext(ContextType.Domain, _myDomainName, adminUsername, adminPassword))                                                               // Check if the the password/user are correct
                {

                    if (context.ConnectedServer != null)                                                                                                                // Throw error if the password/username is incorrect        
                    {

                        isAuthenticated = true;

                        // Publish the credentials before constructing anything, so privileged AD
                        // writes bind as this account rather than as the interactive user.
                        AdminSession.Set(adminUsername, adminPassword, configuration);
                        AppLog.Info($"Connected to Active Directory as: {context.UserName}.", Color.GreenYellow);

                        auditLogManager = new AuditLogManager(adminUsername, configuration);
                        ADGroupManager = new ADGroupActionManager(auditLogManager);
                        PWDManager = new PasswordManager(auditLogManager);
                        ACManager = new AccountCreationManager(auditLogManager, configuration);
                        ADManager.SetAdminCredentials(adminUsername, adminPassword, configuration);

                        bool exit = false;
                        while (!exit)                                                                                                                                          // Loop the menu
                        {
                            DisplayMainMenu();
                            string choice = ConsoleInput.ReadLine();
                            exit = HandleMainMenuChoice(choice, context, ADManager, ADGroupManager, PWDManager, ACManager, ACCDeactivationManager);
                        }// end of while-loop
                    }// end of if statement
                    context.Dispose();
                }// end of using
            }// end of try
            catch (DirectoryServicesCOMException ex) // error out of user credentials are wrong or account is locked
            {
                isAuthenticated = false; // reset so the loop retries credentials
                AppLog.Error("Error: Unable to connect to the Active Directory server. Please check your credentials and try again.", ex);
            }
            catch (Exception ex)
            {
                isAuthenticated = false; // reset so the loop retries credentials
                AppLog.Error($"An error occurred: {ex.Message}", ex, Color.IndianRed);
            }// end of Catch
            // Without this the loop would spin forever on a closed or exhausted stdin, since
            // there is no further input to change the outcome.
            if (ConsoleInput.EndOfInput && !isAuthenticated)
            {
                AppLog.Error("No more console input — giving up on authentication.");
                break;
            }
        } while (!isAuthenticated || string.IsNullOrEmpty(adminUsername));                                                                                                                                            // Repeat until a valid password is entered
        }
        finally
        {
            // Targets are async, so without this the last entries are lost on exit.
            auditLogManager?.EndSession();
            NLog.LogManager.Shutdown();
        }
    }// end of Main Method


    //--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    //                                                                                                          UI
    //--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    static void DisplayMainMenu()
    {
        ConsoleUi.Header();
        ConsoleUi.Menu("Main Menu",
            "Locked Out Management",
            "Group Management",
            "User Information",
            "Reports",
            "Exit");
        ConsoleUi.Prompt("Choice");
    }// end of DisplayMainMenu

    /// <summary>
    /// Main Menu interface
    /// </summary>
    /// <param name="choice"> User choice</param>
    /// <param name="context"> Active directory object</param>
    /// <param name="ADManager"> A class that manage user lockout</param>
    /// <param name="ADGroupManager">A class that manage user groups </param>
    /// <param name="PWDManager"> A class that manager user password related events</param>
    /// <returns></returns>
    static bool HandleMainMenuChoice(string choice, PrincipalContext context, ActiveDirectoryManager ADManager, ADGroupActionManager ADGroupManager, PasswordManager PWDManager, AccountCreationManager ACManager, AccountDeactivationManager ACCDeactivationManager)
    {
        switch (choice)
        {
            case "1":
                DisplayLockedOutMenu(context, ADManager);
                break;
            case "2":
                DisplayGroupManagementMenu(context, ADGroupManager);
                break;
            case "3":
                DisplayUserInfoMenu(context, ADManager, PWDManager, ACManager, ACCDeactivationManager);
                break;
            case "4":
                DisplayReportsMenu(context, ADManager);
                break;
            case "5":
                return true;
            case "clear":
                Console.Clear();
                break;
            default:
                ConsoleUi.Warn("Invalid option. Please try again.");
                break;
        }// end of switch case
        return false;
    }// end of Handle

    /// <summary>
    /// A UI that host all user lockout management
    /// </summary>
    /// <param name="context"></param>
    /// <param name="ADManager"></param>
    static void DisplayLockedOutMenu(PrincipalContext context, ActiveDirectoryManager ADManager)
    {
        bool exit = false;
        while (!exit)
        {
            ConsoleUi.Breadcrumb("Main", "Locked Out Management");
            ConsoleUi.Menu("Locked Out Manager",
                "Unlock a Specific User",
                "Check All Locked Accounts",
                "Unlock All Locked Accounts",
                "Find Lockout Source for a User");
            ConsoleUi.PromptWithExit("Choice");

            string choice = ConsoleInput.ReadTrimmedLower();
            switch (choice)
            {
                case "1":
                    ADManager.UnlockUser(context);
                    break;
                case "2":
                    ADManager.CheckLockedAccounts(context);
                    break;
                case "3":
                    ADManager.UnlockAllUsers(context);
                    break;
                case "4":
                    ADManager.FindLockoutSource();
                    break;
                case "exit":
                    exit = true;
                    break;
                default:
                    ConsoleUi.Warn("Invalid option. Please try again.");
                    break;
            }// end of switch-case
        }// end of while
        Console.Clear();
    }// end of DisplayLockedOutMenu

    /// <summary>
    /// A UI that host all security group and distirbution list management 
    /// </summary>
    /// <param name="context"></param>
    /// <param name="ADGroupManager"></param>
    static void DisplayGroupManagementMenu(PrincipalContext context, ADGroupActionManager ADGroupManager)
    {
        bool exit = false;
        while (!exit)
        {
            ConsoleUi.Breadcrumb("Main", "Group Management");
            ConsoleUi.Menu("Group Management",
                "List All Groups in Active Directory",
                "Add User to a Group",
                "Remove User From a Group",
                "Grant Shared Mailbox Access",
                "Revoke Shared Mailbox Access",
                "Check Who is Member in a Group",
                "Copy Groups From Another User");
            ConsoleUi.PromptWithExit("Choice");

            string choice = ConsoleInput.ReadTrimmedLower();
            switch (choice)
            {
                case "1":
                    ADGroupManager.ListAllGroups(context);
                    break;
                case "2":
                    ADGroupManager.AddUserToGroup(context);
                    break;
                case "3":
                    ADGroupManager.RemoveUserFromGroup(context);
                    break;
                case "4":
                    ADGroupManager.AddUserToSharedMailbox(context);
                    break;
                case "5":
                    ADGroupManager.RemoveUserFromSharedMailbox(context);
                    break;
                case "6":
                    ADGroupManager.ListGroupMembers(context);
                    break;
                case "7":
                    ADGroupManager.CopyGroupsFromUser(context);
                    break;
                case "exit":
                    exit = true;
                    break;
                default:
                    ConsoleUi.Warn("Invalid option. Please try again.");
                    break;
            }// end of switch-case
        }// end of while loop
        Console.Clear();
    }// end of DisplayGroupMangementMenu

    /// <summary>
    /// A UI that host all user info management. 
    /// </summary>
    /// <param name="context"></param>
    /// <param name="ADManager"></param>
    /// <param name="PWDManager"></param>
    static void DisplayUserInfoMenu(PrincipalContext context, ActiveDirectoryManager ADManager, PasswordManager PWDManager, AccountCreationManager ACManager, AccountDeactivationManager ACCDeactivationManager)
    {
        bool exit = false;
        while (!exit)
        {
            ConsoleUi.Breadcrumb("Main", "User Information");
            ConsoleUi.Menu("User Information",
                "Check User Password Expiration Date",
                "Display General User Info",
                "Reset A User Password",
                "Create New User Account",
                "Disable User Account",
                "Find a User (partial name search)");
            ConsoleUi.PromptWithExit("Choice");

            string choice = ConsoleInput.ReadTrimmedLower();
            switch (choice)
            {
                case "1":
                    PWDManager.GetPasswordExpirationDate();
                    break;
                case "2":
                    ADManager.DisplayUserInfo(context);
                    break;
                case "3":
                    PWDManager.ResetUserPassowrd();
                    break;
                case "4":
                    ACManager.CreateUserAccount(adminUsername, adminPassword);
                    break;
                case "5":
                    ACCDeactivationManager.DeactivateUserAccount(context, adminUsername, adminPassword);
                    break;
                case "6":
                    ADManager.FindUsers(context);
                    break;
                case "exit":
                    exit = true;
                    break;
                default:
                    ConsoleUi.Warn("Invalid option. Please try again.");
                    break;
            }// end of switch-case
        }// end of while loop
        Console.Clear();
    }// end of DisplayUserInfoMenu

    /// <summary>
    /// Read-only reports. Grouped separately from the action menus so it is obvious that nothing
    /// here changes anything.
    /// </summary>
    static void DisplayReportsMenu(PrincipalContext context, ActiveDirectoryManager ADManager)
    {
        bool exit = false;
        while (!exit)
        {
            ConsoleUi.Breadcrumb("Main", "Reports");
            ConsoleUi.Menu("Reports (read-only)",
                "Accounts Due for Deletion",
                "Passwords Expiring Soon",
                "Recent Lockouts (all DCs)",
                "Validate Group Assignment (OUs + groups exist)",
                "Compare a Role Against Its Peers");
            ConsoleUi.PromptWithExit("Choice");

            string choice = ConsoleInput.ReadTrimmedLower();
            switch (choice)
            {
                case "1":
                    ADManager.ReportAccountsDueForDeletion(context);
                    break;
                case "2":
                    ADManager.ReportPasswordsExpiringSoon(context);
                    break;
                case "3":
                    ADManager.ReportRecentLockouts();
                    break;
                case "4":
                    ADManager.ReportGroupAssignmentValidation(context);
                    break;
                case "5":
                    ADManager.ReportRoleGroupDrift(context);
                    break;
                case "exit":
                    exit = true;
                    break;
                default:
                    ConsoleUi.Warn("Invalid option. Please try again.");
                    break;
            }// end of switch-case
        }// end of while loop
        Console.Clear();
    }// end of DisplayReportsMenu

}// end of class
