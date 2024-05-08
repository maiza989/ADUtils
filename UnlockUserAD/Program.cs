
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using UnlockUserAD;


// TODO - Add feature to display info of users
// TODO - Add feature to Reset user password 

class Program
{
    static bool isLocked = false;
    static int countdownSeconds = 60;

    void Main(string[] args)
    {
       
        string adminUsername, adminPassword;

        do
        {
            Console.Write("Enter admin username: ");
            adminUsername = Console.ReadLine().Trim();
            Console.Write("Enter admin password: ");
            adminPassword = PasswordManager.GetPassword().Trim();

            try
            {
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain, null, adminUsername, adminPassword))
                {
                  if(context.ConnectedServer != null) 
                  {

                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Connected to Active Directory as: {adminUsername}.");
                    Console.ForegroundColor = ConsoleColor.Gray;

                    bool exit = false;
                    while (!exit)
                    {
                        Console.ForegroundColor = ConsoleColor.Gray;
                        Console.WriteLine("\nSelect an option:");
                        Console.WriteLine("1. Unlock a specific user");
                        Console.WriteLine("2. Check all locked accounts");
                        Console.WriteLine("3. Unlock all locked accounts");
                        Console.WriteLine("4. List all groups in active directory");
                        Console.WriteLine("5. Add User to a group");
                        Console.WriteLine("6. Check user password expiration date");
                        Console.WriteLine("7. Exit");
                        Console.Write("Enter your choice: ");

                        string choice = Console.ReadLine();
                        switch (choice)
                        {
                            case "1":
                                ActiveDirectoryManager.UnlockUser(context);                                                                                                                              // 1 To Unlock Specific user
                                break;
                            case "2":
                                ActiveDirectoryManager.CheckLockedAccounts(context);                                                                                                   // 2 To Check all Unlocked users
                                break;
                            case "3":
                                ActiveDirectoryManager. UnlockAllUsers(context);                                                                                                        // 3 To Unlock all users
                                break;
                            case "4":
                                ActiveDirectoryManager.ListAllGroups(context);                                                                                                         // 4 To list all groups in AD
                                break;
                            case "5":
                                ActiveDirectoryManager.AddUserToGroup(context);                                                                                                        // 5 To add user to a group 
                                break;
                            case "6":
                                PasswordManager.PasswordExpirationDate();                                                                                        // 6 check user password experation date
                                break;
                            case "7":
                                exit = true;                                                                                                                    // 7 To Exit/Close application
                                break;
                            default:
                                Console.WriteLine("Invalid option. Please try again.");
                                break;
                        }// end of switch-case
                    }// end of while-loop
                }// end of using
            }// end of Try-Catch
            }
            catch (DirectoryServicesCOMException)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: Unable to connect to the Active Directory server. Please check your credentials and try again.");
                Console.ForegroundColor = ConsoleColor.Gray;
                adminPassword = ""; // Clear password if incorrect
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of Catch
        } while (string.IsNullOrEmpty(adminPassword)); // Repeat until a valid password is entered
    }// end of Main Method

}// end of class
