
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using UnlockUserAD;
using System.Diagnostics;
using System.Linq;


// TODO - Add feature to display info of users

namespace UnlockUserAD
{
    public class ActiveDirectoryManager
    {
        PasswordManager passwordManager = new PasswordManager();
        public void DisplayUserInfo(PrincipalContext context)
        {
            bool returnToMenu = false;
            List<string> userGroups = new List<string>();
            do
            {
                Console.Write("Enter the username to unlock (type 'exit' to return to the main menu): ");
                string username = Console.ReadLine().Trim().ToLower();

                if (username.ToLower().Trim() == "exit")
                {
                    returnToMenu = true;
                }// end of if statement
                else
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
                    if(user != null)
                    {
                        var groups = user.GetGroups();
                        foreach(var group in groups)
                        {
                            userGroups.Add(group.Name);
                        }
                        string userGroupString = string.Join(", ", userGroups);

                        DirectoryEntry directoryEntry = user.GetUnderlyingObject() as DirectoryEntry;
                        string title = directoryEntry.Properties["title"].Value as string;

                        Console.WriteLine($"\n\tFirst name: {user.GivenName}\n" +
                                          $"\tLast name: {user.Surname}\n" +
                                          $"\tDisplay name: {user.DisplayName}\n" +
                                          $"\tUsername: {user.SamAccountName}\n" +
                                          $"\tEmail: {user.EmailAddress}\n" +
                                          $"\tTitle: {title}\n" +
                                          $"\tMember of: {userGroupString}\n" +
                                          $"\tPassword Last Set: {passwordManager.GetPasswordLastSetDate(user)}\n" +
                                          $"\tPassword Experation Date: {passwordManager.GetPasswordExpirationDate(user)}\n" +
                                          $"\tBad Logon Counter: {user.BadLogonCount}\n" +
                                          $"\tLast Bad Logon Attempt: {user.LastBadPasswordAttempt}\n" +
                                          $"\tAccount Status: {user.Enabled}\n" +
                                          $"\tAccount Lockout Status: {user.IsAccountLockedOut()}\n" +
                                          $"\tHome Directory: {user.HomeDirectory}\n" +
                                          $"\tSID: {user.Sid}\n" +
                                          $"");
                    }
                }
            }while (!returnToMenu);
        }
        /// <summary>
        /// A method to unlock one sepcific user.
        /// </summary>
        /// <param name="context">Based in what the computer domain</param>
        public void UnlockUser(PrincipalContext context)                                                                                      
        {
            bool returnToMenu = false;
            do
            {
                Console.Write("Enter the username to unlock (type 'menu' to return to the main menu): ");
                string username = Console.ReadLine().Trim();

                if (username.ToLower() == "menu")
                {
                    returnToMenu = true;
                }// end of if statement
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
                                
                            }// end of inner-if-statement
                            else
                            {
                                Console.ForegroundColor = ConsoleColor.DarkRed;
                                Console.WriteLine($"\tUser account '{username}' is not locked.");
                                Console.ForegroundColor = ConsoleColor.Gray;
                            }// end of else-statement
                        }// end of Outter-if-statement
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.DarkRed;
                            Console.WriteLine($"\tUser account '{username}' not found in Active Directory.");
                            Console.ForegroundColor = ConsoleColor.Gray;

                        }// end of else-statement
                    }// end of Try-Catch
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Error Unlocking a user: {ex.Message}");
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }// end of catch
                }// end of else
            } while (!returnToMenu);
        }// end of UnlockUser

        /// <summary>
        /// A method to go through every user in Active Directory and unlock all of them if any is locked.
        /// </summary>
        /// <param name="context">Based in what the computer domain</param>
        public void UnlockAllUsers(PrincipalContext context)                                                                           
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

        // TODO - grab the lockout event
       
        /// <summary>
        /// A method to check if any user is locked in Active Directory.
        /// </summary>
        /// <param name="context">Based in what the computer domain</param>
        public void CheckLockedAccounts(PrincipalContext context)                                                                        
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
    }// end of class
}// end of spacename
