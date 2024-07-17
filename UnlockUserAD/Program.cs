
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Security.Cryptography.X509Certificates;
using UnlockUserAD;
// TODO - DONE - Audit Logging: Log fuction to record important action performed.

class Program
{
    static bool isLocked = false;
    static int countdownSeconds = 60;
    static string adminUsername, adminPassword;
    static private bool isAuthenticated = false;

    static void GetAdminCreditials()
    {
        Console.Write("Enter admin username: ");
        adminUsername = Console.ReadLine().Trim();
        Console.Write("Enter admin password: ");
        adminPassword = PasswordManager.GetPassword().Trim();
    }
    static void Main(string[] args)
    {
        ActiveDirectoryManager ADManager = new ActiveDirectoryManager();
        AccountCreationManager ACManager;
        PasswordManager PWDManager = null;
        ADGroupActionManager ADGroupManager = null;
        AuditLogManager auditLogManager = null;
        
        do
        {
        GetAdminCreditials();
            try
            {
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain, null, adminUsername, adminPassword))                                               // Check if the the password/user are correct
                {
                    if (context.ConnectedServer != null)                                                                                                                      // Throw error if the password/username is incorrect        
                    {
                        isAuthenticated = true;

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Connected to Active Directory as: {adminUsername}.");
                        Console.ForegroundColor = ConsoleColor.Gray;

                        auditLogManager = new AuditLogManager(adminUsername);
                        ADGroupManager = new ADGroupActionManager(auditLogManager);
                        PWDManager = new PasswordManager(auditLogManager);
                        ACManager = new AccountCreationManager(auditLogManager);


                        bool exit = false;
                        while (!exit)                                                                                                                                          // Loop the menu
                        {
                            DisplayMainMenu();
                            string choice = Console.ReadLine();
                            exit = HandleMainMenuChoice(choice, context, ADManager, ADGroupManager, PWDManager, ACManager);
                        }// end of while-loop
                    }// end of if statement
                    context.Dispose();
                }// end of using
            }// end of Try-Catch
            catch (DirectoryServicesCOMException)                                                                                                                              // Error out if password/username are incorrect
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: Unable to connect to the Active Directory server. Please check your credentials and try again.");
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of Catch
        } while (!isAuthenticated && string.IsNullOrEmpty(adminUsername));                                                                                                                                            // Repeat until a valid password is entered
    }// end of Main Method


    //--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    //                                                                                                          UI
    //--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    static void DisplayMainMenu()
    {

        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine("\nSelect an option:");
        Console.WriteLine("1. Locked Out Management");
        Console.WriteLine("2. Group Management");
        Console.WriteLine("3. User Information Management");
        Console.WriteLine("4. Exit");
        Console.Write("Enter your choice: ");
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
    static bool HandleMainMenuChoice(string choice, PrincipalContext context, ActiveDirectoryManager ADManager, ADGroupActionManager ADGroupManager, PasswordManager PWDManager, AccountCreationManager ACManager)
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
                DisplayUserInfoMenu(context, ADManager, PWDManager, ACManager);
                break;
            case "4":
                return true;
            case "clear":
                Console.Clear();
                break;
            default:
                Console.WriteLine("Invalid option. Please try again.");
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
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("\nLocked Out Manager:");
            Console.WriteLine("1. Unlock a Specific User");
            Console.WriteLine("2. Check All Locked Accounts");
            Console.WriteLine("3. Unlock All Locked Accounts");
            Console.Write("Enter your choice(Type \"exit\" to return to main menu): ");
            string choice = Console.ReadLine().ToLower().Trim();
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
                case "exit":
                    exit = true;
                    break;
                default:
                    Console.WriteLine("Invalid option. Please try again.");
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
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("\nGroup Management:");
            Console.WriteLine("1. List All Groups in Active Directory");
            Console.WriteLine("2. Add User to a Group");
            Console.WriteLine("3. Remove User From a Group");
            Console.WriteLine("4. Check Who is Member in a Group");
            Console.Write("Enter your choice(Type \"exit\" to return to main menu): ");

            string choice = Console.ReadLine().ToLower().Trim();
            switch (choice)
            {
                case "1":
                    ADGroupManager.ListAllGroups(context);
                    break;
                case "2":
                    ADGroupManager.AddUserToGroup(context);
                    break;
                case "3":
                    ADGroupManager.RemoveUserToGroup(context);
                    break;
                case "4":
                    ADGroupManager.ListGroupMembers(context);
                    break;
                case "exit":
                    exit = true;
                    break;
                default:
                    Console.WriteLine("Invalid option. Please try again.");
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
    static void DisplayUserInfoMenu(PrincipalContext context, ActiveDirectoryManager ADManager, PasswordManager PWDManager, AccountCreationManager ACManager)
    {
        bool exit = false;
        while (!exit)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("\nUser Information:");
            Console.WriteLine("1. Check User Password Expiration Date");
            Console.WriteLine("2. Display General User Info");
            Console.WriteLine("3. Reset A User Password");
            Console.WriteLine("4. Create New User Account");
            Console.Write("Enter your choice(Type \"exit\" to return to main menu): ");

            string choice = Console.ReadLine().ToLower().Trim();
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
                case "exit":
                    exit = true;
                    break;
                default:
                    Console.WriteLine("Invalid option. Please try again.");
                    break;
            }// end of switch-case
        }// end of while loop
        Console.Clear();
    }// end of DisplayUserInfoMenu

    static void DisplayUserCreationMenu()
    {

    }
}// end of class
