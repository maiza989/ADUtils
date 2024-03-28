using System;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Threading;

class Program
{
    static bool isLocked = false;
    static int countdownSeconds = 60;
    static void Main(string[] args)
    {
        Console.Write("Enter admin username: ");                                                                                                          // Prompt for admin credentials
        string adminUsername = Console.ReadLine().Trim();
        Console.Write("Enter admin password: ");
        string adminPassword = GetPassword().Trim();

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
                    Console.WriteLine("4. List all groups in active directory");
                    Console.WriteLine("5. Add User to a group");
                    Console.WriteLine("6. Check user password expiration date **PENDING**");
                    Console.WriteLine("7. Exit");
                    Console.Write("Enter your choice: ");

                    string choice = Console.ReadLine();
                    switch (choice)
                    {
                        case "1":
                            UnlockUser(context);                                                                                                            // 1 To Unlock Specific user
                            break;
                        case "2":
                            CheckLockedAccounts(context);                                                                                                   // 2 To Check all Unlocked users
                            break;
                        case "3":
                            UnlockAllUsers(context);                                                                                                        // 3 To Unlock all users
                            break;
                        case "4":
                            ListAllGroups(context);                                                                                                         // 4 To list all groups in AD
                            break;
                        case "5":
                            AddUserToGroup(context);                                                                                                        // 5 To add user to a group 
                            break;
                        case "6":
                            CheckExpiredPassword(context);
                            break;
                        case "7":
                            exit = true;                                                                                                                    // 6 To Exit/Close application
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
            string username = Console.ReadLine().Trim();

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

    static List<string> QuerySecurityLogs(string username)
    {
        List<string> lockoutEvents = new List<string>();
        try
        {
            EventLog eventLog = new EventLog("Security");

            // Filter for Event ID 4740 (Account Lockout)
            var lockoutEntries = eventLog.Entries.Cast<EventLogEntry>()
                .Where(entry => entry.InstanceId == 4656 && entry.Message.Contains(username));

            foreach (var entry in lockoutEntries)
            {
                lockoutEvents.Add(entry.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error querying security logs: {ex.Message}");
        }
        return lockoutEvents;
    }// end of QuerySecurityLogs
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

                   /* List<string> lockoutEvents = QuerySecurityLogs(user.SamAccountName);
                    if (lockoutEvents.Any())
                    {
                        Console.WriteLine("\t  Lockout events:");
                        foreach (var evt in lockoutEvents)
                        {
                            Console.WriteLine($"\t    - {evt}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("\t  No lockout events found in security logs.");
                    }*/
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

    static void AddUserToGroup(PrincipalContext context)
    {
        bool isExit = false;
        do
        {
            Console.Write("Enter the username(Type 'exit' to go back to menu): ");
            string username = Console.ReadLine().Trim();
            if(username.ToLower().Trim() == "exit")
            {
                isExit = true;
                Console.WriteLine($"\nReturing to menu...");
                break;
            }
            Console.Write("Enter the group name (Type 'exit' to go back to menu): ");
            string groupName = Console.ReadLine().Trim();  
            if(groupName.ToLower().Trim() == "exit")
            {
                isExit = true;
                Console.WriteLine($"Returing to menu...");
                break;
            }
            else
            {

                try
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);                                              // Check for user in AD

                    if (user != null)
                    {

                        GroupPrincipal group = GroupPrincipal.FindByIdentity(context, groupName);                                                                   // Check for group in AD

                        if (group != null)
                        {
                            if (!group.Members.Contains(user))                                                                                                      // If the user is not in the group add him
                            {
                                group.Members.Add(user);                                                                                                            // Add the user to the group
                                group.Save();                                                                                                                       // Apply changes
                                group.Dispose();
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine($"User '{username}' added to group '{groupName}' successfully.");
                                Console.ForegroundColor = ConsoleColor.Gray;
                            }// end of inner-2 if-statement
                            else
                            {
                                Console.ForegroundColor = ConsoleColor.DarkYellow;
                                Console.WriteLine($"User '{username}' is already a member of group '{groupName}'.");
                                Console.ForegroundColor = ConsoleColor.Gray;
                            }// end of inner-2 else-statement
                        }// end of inner if-statement
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.DarkRed;
                            Console.WriteLine($"Group '{groupName}' not found in Active Directory.");
                            Console.ForegroundColor = ConsoleColor.Gray;
                        }// end of outter else-statement
                    }// end of outter if-statement 
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine($"User '{username}' not found in Active Directory.");
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }
                }// end of try
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Error adding user to group: {ex.Message}");
                    Console.ForegroundColor = ConsoleColor.Gray;
                }// end of catch
            }    
        } while(!isExit);
    }// end of AddUserToGroup
    static void ListAllGroups(PrincipalContext context)
    {
        Console.WriteLine("\nList of all groups:");

        try
        {
            Console.Write("Enter the first letter of the group name to filter (or press Enter to show all groups): ");
            char filterLetter = Console.ReadKey().KeyChar;
            Console.WriteLine();

            PrincipalSearcher searcher = new PrincipalSearcher(new GroupPrincipal(context));                                                                    // Search for all groups
            List<string> groupNames = new List<string>();

            foreach (var result in searcher.FindAll())
            {
                GroupPrincipal group = result as GroupPrincipal;
                if (char.ToLower(group.Name[0]) == char.ToLower(filterLetter) || filterLetter == '\r')                                                          // Filter by the first letter or show all groups if Enter is pressed)
                {
                    groupNames.Add(group.Name);
                }
            }// end of foreach

            groupNames.Sort();
            int maxGroupNameLength = groupNames.Max(g => g.Length);                                                                                             // use max method to find the longest group length
            int columnWidth = maxGroupNameLength + 5;                                                                                                           // Add padding

            int numColumns = Console.WindowWidth / columnWidth;                                                                                                 // Calculate number of columns based on window width
            int numRows = (int)Math.Ceiling((double)groupNames.Count / numColumns);                                                                             // Calculate number of rows. If total number of gorups does not evenly fit into the columns, Matha.Ceiling rounds it to the nearst integer.

            for (int i = 0; i < numRows; i++)                                                                                                                   // Nested for loop to print Group names in a grid style
            {
                for (int j = 0; j < numColumns; j++)
                {
                    int index = i + j * numRows;                                                                                                                // Calculate the index of the gorup base on 'i' rows and 'j' columns
                    if (index < groupNames.Count)
                    {
                        Console.Write($"- {groupNames[index].PadRight(columnWidth)}");                                                                          // Print each group name with specified right padding
                    }
                }// end of inner for loop
                Console.WriteLine();
            }// end of outter for loop
        }// end of try
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error listing groups: {ex.Message}");
            Console.ForegroundColor = ConsoleColor.Gray;
        }// end of catch
    }// end of ListAllGroups

    static void CheckExpiredPassword(PrincipalContext context)
    {
        Console.Write("Enter the username to check password expiration: ");
        string username = Console.ReadLine();

        try
        {
            UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);

            if (user != null)
            {
                if (user.PasswordNeverExpires) // Check if the password never expires
                {
                        
                    DateTime passwordExpiration = user.AccountExpirationDate.Value.ToLocalTime();

                    if (passwordExpiration < DateTime.Today)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"User '{username}' password has expired.");
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"User '{username}' password is not expired.");
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine($"Password for user '{username}' never expires.");
                    Console.ForegroundColor = ConsoleColor.Gray;
                }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine($"User '{username}' not found in Active Directory.");
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error checking password expiration: {ex.Message}");
            Console.ForegroundColor = ConsoleColor.Gray;
        }
    }


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
