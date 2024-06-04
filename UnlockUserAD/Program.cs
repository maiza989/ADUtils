
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Security.Cryptography.X509Certificates;
using UnlockUserAD;



// TODO - User Account Creation: Enable users to create new accounts in Active Directory. 
// TODO - Audit Logging: Log fuction to record important action performed.



class Program
{
    static bool isLocked = false;
    static int countdownSeconds = 60;

    static void Main(string[] args)
    {
        ActiveDirectoryManager ADManager = new ActiveDirectoryManager();
        ADGroupActionManager ADGroupManager = new ADGroupActionManager();
        PasswordManager PWDManager = new PasswordManager();
        string adminUsername, adminPassword;
        
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
                        while (!exit)                                                                                                                                        // Loop the menu
                        {
                            Console.ForegroundColor = ConsoleColor.Gray;
                            Console.WriteLine("\nSelect an option:");
                            Console.WriteLine("1. Unlock a Specific User");
                            Console.WriteLine("2. Check All Locked Accounts");
                            Console.WriteLine("3. Unlock All Locked Accounts");
                            Console.WriteLine("4. List All Groups in Active Directory");
                            Console.WriteLine("5. Add User to a group");
                            Console.WriteLine("6. Remove User From a Group");
                            Console.WriteLine("7. Check Who is Memeber in a Group");
                            Console.WriteLine("8. Check User Password Expiration Date");
                            Console.WriteLine("9. Exit");
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
                                    ADManager. UnlockAllUsers(context);                                                                                                       // 3 To Unlock all users
                                    break;
                                case "4":
                                    ADGroupManager.ListAllGroups(context);                                                                                                    // 4 To list all groups in AD
                                    break;
                                case "5":
                                    ADGroupManager.AddUserToGroup(context);                                                                                                   // 5 To add user to a group 
                                    break;
                                case "6":
                                    ADGroupManager.RemoveUserToGroup(context);
                                    break;
                                case "7":
                                    ADGroupManager.ListGroupMembers(context);
                                    break;
                                case "8":
                                    PWDManager.GetPasswordExpirationDate();                                                                                                   // 6 check user password experation date
                                    break;
                                case "9":
                                    exit = true;                                                                                                                              // 9 To Exit/Close application
                                    break;
                                default:
                                    Console.WriteLine("Invalid option. Please try again.");
                                    break;
                            }// end of switch-case
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

}// end of class
