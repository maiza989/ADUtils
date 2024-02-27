using System;
using System.DirectoryServices.AccountManagement;

class Program
{
    static void Main(string[] args)
    {
        
                                                                                                                                                        // Prompt for admin credentials
        Console.Write("Enter admin username: ");
        string adminUsername = Console.ReadLine();
        Console.Write("Enter admin password: ");
        string adminPassword = GetPassword();

        try
        {
            using (PrincipalContext context = new PrincipalContext(ContextType.Domain, null, adminUsername, adminPassword))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Connected to Active Directory as: {adminUsername}.");

                while (true)
                {
                    Console.ForegroundColor = ConsoleColor.Gray;
                                                                                                                                                       // Prompt for the username to unlock
                    Console.Write("\nEnter the username to unlock (type 'exit' to quit): ");
                    string username = Console.ReadLine();

                                                                                                                                                       // Check if the user wants to exit
                    if (username.ToLower() == "exit")
                        break;

                                                                                                                                                       // Check if the user exists in Active Directory
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
                    if (user != null)
                    {
                                                                                                                                                       // Check if the user account is locked
                        if (user.IsAccountLockedOut())
                        {
                                                                                                                                                       // Unlock the user account
                            user.UnlockAccount();
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine($"\tUser account '{username}' has been unlocked.");
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.DarkRed;
                            Console.WriteLine($"\tUser account '{username}' is not locked.");
                        }
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine($"\tUser account '{username}' not found in Active Directory.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"An error occurred: {ex.Message}");
        }

        Console.WriteLine("Press any key to exit.");
        Console.ReadKey();
    }
                                                                                                                                        // Method to read password without displaying it on the console
    static string GetPassword()
    {
        string password = "";
        ConsoleKeyInfo key;
        do
        {
            key = Console.ReadKey(true);

                                                                                                                                        // Ignore any key that isn't a printable character
            if (!char.IsControl(key.KeyChar))
            {
                password += key.KeyChar;
                Console.Write("*");
            }
            else if (key.Key == ConsoleKey.Backspace && password.Length > 0)
            {
                password = password.Substring(0, password.Length - 1);
                Console.Write("\b \b");
            }
        } while (key.Key != ConsoleKey.Enter);

        Console.WriteLine();
        return password;
    }
}

