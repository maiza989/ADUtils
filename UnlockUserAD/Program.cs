
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Security.Cryptography.X509Certificates;
using UnlockUserAD;




// TODO - Audit Logging: Log fuction to record important action performed.


class Program
{
    static bool isLocked = false;
    static int countdownSeconds = 60;
    static string adminUsername, adminPassword;
    static void Main(string[] args)
    {
        PasswordManager PWDManager = new PasswordManager();
        ActiveDirectoryManager ADManager = new ActiveDirectoryManager();
        ADGroupActionManager ADGroupManager = new ADGroupActionManager();
        
        do
        {
            Console.Write("Enter admin username: ");
            adminUsername = Console.ReadLine().Trim();
            Console.Write("Enter admin password: ");
            adminPassword = PasswordManager.GetPassword().Trim();
            
            try
            {
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain, null, adminUsername, adminPassword))                                              // Check if the the password/user are correct
                {
                    if(context.ConnectedServer != null)                                                                                                                      // Throw error if the password/username is incorrect        
                    {

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Connected to Active Directory as: {adminUsername}.");
                        Console.ForegroundColor = ConsoleColor.Gray;

                        bool exit = false;
                        while (!exit)                                                                                                                                          // Loop the menu
                        {
                              DisplayMainMenu();
                              string choice = Console.ReadLine();
                              exit = HandleMainMenuChoice(choice, context, ADManager, ADGroupManager, PWDManager);
                           /* Console.ForegroundColor = ConsoleColor.Gray;
                            Console.WriteLine("\nSelect an option:");
                            Console.WriteLine("1. Unlock a Specific User");
                            Console.WriteLine("2. Check All Locked Accounts");
                            Console.WriteLine("3. Unlock All Locked Accounts");
                            Console.WriteLine("4. Check User Password Expiration Date");
                            Console.WriteLine("5. List All Groups in Active Directory");
                            Console.WriteLine("6. Add User to a group");
                            Console.WriteLine("7. Remove User From a Group");
                            Console.WriteLine("8. Check Who is Memeber in a Group");
                            Console.WriteLine("9. Display General A User Info");
                            Console.WriteLine("10. Exit");
                            Console.Write("Enter your choice: ");

                            string choice = Console.ReadLine();
                            switch (choice)
                            {
                                case "1":
                                    ADManager.UnlockUser(context);                                                                                                            // 1 To Unlock Specific user
                                    break;
                                case "2":
                                    ADManager.CheckLockedAccounts(context);                                                                                                   // 2 To Check all Unlocked users
                                    break;
                                case "3":
                                    ADManager.UnlockAllUsers(context);                                                                                                       // 3 To Unlock all users
                                    break;
                                case "4":
                                    ADGroupManager.ListGroupMembers(context);
                                    break;
                                case "5":
                                    ADGroupManager.ListAllGroups(context);                                                                                                    // 4 To list all groups in AD
                                    break;
                                case "6":
                                    ADGroupManager.AddUserToGroup(context);                                                                                                   // 5 To add user to a group 
                                    break;
                                case "7":
                                    ADGroupManager.RemoveUserToGroup(context);
                                    break;
                                case "8":
                                    PWDManager.GetPasswordExpirationDate();                                                                                                   // 6 check user password experation date
                                    break;
                                case "9":
                                    ADManager.DisplayUserInfo(context);
                                    break;
                                case "10":
                                    exit = true;                                                                                                                              // 9 To Exit/Close application
                                    break;
                                default:
                                    Console.WriteLine("Invalid option. Please try again.");
                                    break;
                            }// end of switch-case*/
                        }// end of while-loop
                    }// end of if statement
                }// end of using
            }// end of Try-Catch
            catch (DirectoryServicesCOMException)                                                                                                                              // Error out if password/username are incorrect
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: Unable to connect to the Active Directory server. Please check your credentials and try again.");
                Console.ForegroundColor = ConsoleColor.Gray;
                adminPassword = "";                                                                                                                                            // Clear password if incorrect
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of Catch
        } while (string.IsNullOrEmpty(adminPassword));                                                                                                                         // Repeat until a valid password is entered
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
    static bool HandleMainMenuChoice(string choice, PrincipalContext context, ActiveDirectoryManager ADManager, ADGroupActionManager ADGroupManager, PasswordManager PWDManager)
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
                DisplayUserInfoMenu(context, ADManager, PWDManager);
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
    static void DisplayUserInfoMenu(PrincipalContext context, ActiveDirectoryManager ADManager, PasswordManager PWDManager)
    {
        bool exit = false;
        while (!exit)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("\nUser Information:");
            Console.WriteLine("1. Check User Password Expiration Date");
            Console.WriteLine("2. Display General User Info");
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
