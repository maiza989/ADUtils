
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using ADUtils;
using System.Diagnostics;
using System.Linq;
using Pastel;
using System.Drawing;

namespace ADUtils
{

    public class ActiveDirectoryManager
    {
        PasswordManager passwordManager = new PasswordManager();


        /// <summary>
        /// A method that display a general information about a user.
        /// </summary>
        /// <param name="context">The PrincipalContext to use for querying Active Directory</param>
        public void DisplayUserInfo(PrincipalContext context)
        {
            bool returnToMenu = false;
            List<string> userGroups = new List<string>();
            do
            {
                Console.Write($"Enter the username to display info about (Type {"'exit'".Pastel(Color.MediumPurple)} to return to the main menu): ");
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
                        }// end of foreach
                        string userGroupsString = string.Join(", ", userGroups);

                        DirectoryEntry directoryEntry = user.GetUnderlyingObject() as DirectoryEntry;
                        string title = directoryEntry.Properties["title"].Value as string;
                        string department = directoryEntry.Properties["department"].Value as string;
                 
                        DateTime lastBadPasswordAttemptLocal = TimeZoneInfo.ConvertTimeFromUtc(user.LastBadPasswordAttempt.Value.ToUniversalTime(), TimeZoneInfo.Local);
                        DateTime lastLogonLocal = TimeZoneInfo.ConvertTimeFromUtc(user.LastLogon.Value.ToUniversalTime(), TimeZoneInfo.Local);
                        
                        Console.WriteLine($"\nFirst name: {user.GivenName ?? "N/A"}\n" +
                                          $"Last name: {user.Surname ?? "N/A"}\n" +
                                          $"Display name: {user.DisplayName ?? "N/A"}\n" +
                                          $"Username: {user.SamAccountName ?? "N/A"}\n" +
                                          $"Email: {user.EmailAddress ?? "N/A"}\n" +
                                          $"Title: {title ?? "N/A"}\n" +
                                          $"Department: {department ?? "N/A"}\n" +
                                          $"Member of: {userGroupsString ?? "N/A"}\n" +
                                          $"Password Last Set: {passwordManager.GetPasswordLastSetDate(user)}\n" +
                                          $"Password Experation Date: {passwordManager.GetPasswordExpirationDate(user)}\n" +
                                          $"Bad Logon Counter: {user.BadLogonCount}\n" +
                                          $"Last Logon: {lastLogonLocal}\n" +
                                          $"Last Bad Logon Attempt: {lastBadPasswordAttemptLocal}\n" +
                                          $"Account Status: {user.Enabled}\n" +
                                          $"Account Lockout Status: {user.IsAccountLockedOut()}\n" +
                                          $"Home Directory: {user.HomeDirectory ?? "N/A"}\n" +
                                          $"SID: {user.Sid}\n" +
                                          $"");
                                          
                    }// end of if statement
                }// end of else statement
                userGroups.Clear();
            }while (!returnToMenu);
        }// end of DisplayUserInfo
        /// <summary>
        /// A method to unlock one sepcific user.
        /// </summary>
        /// <param name="context">Based in what the computer domain</param>
        public void UnlockUser(PrincipalContext context)                                                                                      
        {
            bool returnToMenu = false;
            do
            {
                Console.Write($"Enter the username to unlock (type {"'exit'".Pastel(Color.MediumPurple)} to return to the main menu): ");
                string username = Console.ReadLine().Trim().ToLower();

                if (username.ToLower().Trim() == "exit")
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
                                Console.WriteLine($"\tUser account '{username}' is not locked.".Pastel(Color.OrangeRed));
                            }// end of else-statement
                        }// end of Outter-if-statement
                        else
                        {
                            Console.WriteLine($"\tUser account '{username}' not found in Active Directory.".Pastel(Color.IndianRed));
                        }// end of else-statement
                    }// end of Try-Catch
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error Unlocking a user: {ex.Message}".Pastel(Color.IndianRed));
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
                        Console.WriteLine($"\tUser account '{user.SamAccountName}' has been unlocked.".Pastel(Color.LimeGreen));
                        anyUnlocked = true;
                    }// end of if-statement
                }// end of foreach
                if (!anyUnlocked)                                                                                                              // If-Else statement to check if any user were unlocked and print appropriate response.
                {
                    Console.WriteLine("\tNo user accounts were locked.".Pastel(Color.DarkGoldenrod));
                }// end of if-statement
                else
                {  
                    Console.WriteLine("\nAll user accounts have been unlocked successfully.".Pastel(Color.DarkCyan));
                }// end of else-statement
            }// end of Try-Catch
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}".Pastel(Color.IndianRed));
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
                        Console.WriteLine($"\t- {user.SamAccountName}".Pastel(Color.Crimson));
                        isAnyLocked = true;
                    }// end of if-statement
                }// end of foreach
                if (!isAnyLocked)
                {
                    Console.WriteLine($"\tNo accounts are LOCKED!!! YAY!!!.".Pastel(Color.RoyalBlue));
                }// end of if-statement
            }// end of try-catch
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}".Pastel(Color.IndianRed));
            }// end of catch
        }// end of CheckLockedAccounts
    }// end of class
}// end of spacename
