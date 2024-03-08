using System;
using System.DirectoryServices.AccountManagement;
using System.Threading;

class Program
{
    static bool isLocked = false;
    static int countdownSeconds = 60;
    static void Main(string[] args)
    {
        Console.Write("Enter admin username: ");                                                                                                          // Prompt for admin credentials
        string adminUsername = Console.ReadLine();
        Console.Write("Enter admin password: ");
        string adminPassword = GetPassword();

        try
        {
            using (PrincipalContext context = new PrincipalContext(ContextType.Domain, null, adminUsername, adminPassword))                               // Login into Active Directory with desire user
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Connected to Active Directory as: {adminUsername}.");
                Console.ForegroundColor = ConsoleColor.Gray;

                bool exit = false;
                while (!exit)                                                                                                                             // Loop Menu 
                {
                    Console.ForegroundColor = ConsoleColor.Gray;
                    Console.WriteLine("\nSelect an option:");
                    Console.WriteLine("1. Unlock a specific user");
                    Console.WriteLine("2. Check all locked accounts");
                    Console.WriteLine("3. Unlock all locked accounts");
                    Console.WriteLine("4. Exit");
                    Console.Write("Enter your choice: ");

                    string choice = Console.ReadLine();
                    switch (choice)
                    {
                        case "1":
                            UnlockUser(context);                                                                                                          // 1 To Unlock Specific user
                            break;
                        case "2":
                            CheckLockedAccounts(context);                                                                                                 // 2 To Check all Unlocked users
                            break;
                        case "3":
                            UnlockAllUsers(context);                                                                                                      // 3 To Unlock all users
                            break;
                        case "4":
                            exit = true;                                                                                                                  // 4 To Exit/Close application
                            break;
                        default:
                            Console.WriteLine("Invalid option. Please try again.");
                            break;
                    }// end of switch-case
                }// end of while-loop
            }// end of using
        }// end of Try-Catch
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"An error occurred: {ex.Message}");
            Console.ForegroundColor = ConsoleColor.Gray;
        }// end of Catch
    }// end of Main Method
    
    static void UnlockUser(PrincipalContext context)                                                                                             // Method to unlock a specific user
    {
        bool returnToMenu = false;
        do
        {
            Console.Write("Enter the username to unlock (type 'menu' to return to the main menu): ");

            string username = Console.ReadLine();
            if (username.ToLower() == "menu")
            {
                returnToMenu = true;
            }
            else
            {
                try
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);                           // Search for specific user using username
                    if (user != null)
                    {
                        if (user.IsAccountLockedOut())                                                                                           // Check if the user is locked
                        {
                            user.UnlockAccount();                                                                                                // Unlock user
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine($"\tUser account '{username}' has been unlocked.");
                            Console.ForegroundColor = ConsoleColor.Gray;
                        }// end of in-if-statement
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.DarkRed;
                            Console.WriteLine($"\tUser account '{username}' is not locked.");
                            Console.ForegroundColor = ConsoleColor.Gray;
                        }// end of else
                    }// end of Out-if-statement
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine($"\tUser account '{username}' not found in Active Directory.");
                        Console.ForegroundColor = ConsoleColor.Gray;

                    }// end of else
                }// end of Try-Catch
                catch(Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Error Unlocking a user: {ex.Message}");
                    Console.ForegroundColor = ConsoleColor.Gray;
                }// end of catch
            }// end of else
        } while (!returnToMenu);
    }// end of UnlockUser

    static void UnlockAllUsers(PrincipalContext context)                                                                                    // Method to unlock all locked accounts
    {
        try
        {
            Console.WriteLine("\nUnlocking all user accounts...");
            PrincipalSearcher searcher = new PrincipalSearcher(new UserPrincipal(context) { Enabled = true });
            bool anyUnlocked = false;
            foreach (var result in searcher.FindAll())
            {
                UserPrincipal user = result as UserPrincipal;
                if (user != null && user.IsAccountLockedOut())                                                                             // If-statement to unlock all users
                {
                    user.UnlockAccount();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"\tUser account '{user.SamAccountName}' has been unlocked.");
                    Console.ForegroundColor = ConsoleColor.Gray;
                    anyUnlocked = true;
                }// end of if-statement
            }// end of foreach
            if (!anyUnlocked)                                                                                                              // If-Else statement to check if any user were unlocked and print appropriate response.
            {
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine("\tNo user accounts were locked.");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of if-statement
            else
            {
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine("\nAll user accounts have been unlocked successfully.");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of else-statement
        }// end of Try-Catch
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error: {ex.Message}");
            Console.ForegroundColor = ConsoleColor.Gray;
        }// end of Catch
    }// end of UnlockAllUsers

    static void CheckLockedAccounts(PrincipalContext context)                                                                               // Method to check for locked accounts
    {
        Console.WriteLine("\nLocked user accounts:");
        try
        {
            PrincipalSearcher searcher = new PrincipalSearcher(new UserPrincipal(context) { Enabled = true });                              // Creating the search object
            bool isAnyLocked = false;
            foreach (var result in searcher.FindAll())                                                                                      // Look through what is in the user search object
            {
                UserPrincipal user = result as UserPrincipal;
                if (user != null && user.IsAccountLockedOut())                                                                              // Print out all locked users
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"\t- {user.SamAccountName}");
                    Console.ForegroundColor = ConsoleColor.Gray;
                    isAnyLocked = true;
                }// end of if-statement
            }// end of foreach
            if (!isAnyLocked)
            {
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine($"\tNo accounts are LOCKED!!! YAY!!!.");
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of if-statement
        }// end of try-catch
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error: {ex.Message}");
            Console.ForegroundColor = ConsoleColor.Gray;
        }// end of catch
    }// end of CheckLockedAccounts
    
    static string GetPassword()                                                                                                           // Method to read password without displaying it on the console
    {
        string password = "";
        ConsoleKeyInfo key;
        do
        {
            key = Console.ReadKey(true);
            if (!char.IsControl(key.KeyChar))                                                                                             // Any key writing will be hiden with * and ignore any key that isn't a printable character
            {
                password += key.KeyChar;
                Console.Write("*");
            }// end of if-statement
            else if (key.Key == ConsoleKey.Backspace && password.Length > 0)                                                              // Give the user ability to backspace
            {
                password = password.Substring(0, password.Length - 1);
                Console.Write("\b \b");
            }// end of else-if
        } while (key.Key != ConsoleKey.Enter);

        Console.WriteLine();
        return password;
    }// end of GetPassword
}// end of class
