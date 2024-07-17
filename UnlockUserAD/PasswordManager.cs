
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Text.RegularExpressions;
using UnlockUserAD;


// TODO - DONE Add feature to Reset user password.

namespace UnlockUserAD
{
    public class PasswordManager
    {
        AuditLogManager auditLogManager;
        DateTime todayDate = DateTime.Now;

        public PasswordManager() { }
        public PasswordManager(AuditLogManager auditLogManager)
        {
           this.auditLogManager = auditLogManager;
        }

        /// <summary>
        /// A method used to reset a user password.
        /// </summary>
        public void ResetUserPassowrd()
        {
            Console.Write("Enter the username to reset password for: ");
            string username = Console.ReadLine();

            try
            {
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);                                           // Searching for the user in AD

                    if (user != null)
                    {
                        DateTime expirationDate = GetPasswordExpirationDate(user);

                        Console.WriteLine("\nPassword Requirement: 15 Characters, Symbols, Number, Lower and upper case. ");
                        Console.Write("Enter the desired password: ");
                        string password = Console.ReadLine();

                        if (IsPasswordVaild(password))
                        {
                            Console.Write($"Password: {password}\n" +
                                              $"Is this the password you desire?(Y/N)");
                            string comfirmation = Console.ReadLine().ToUpper().Trim();

                            if (comfirmation == "Y")
                            {

                                user.SetPassword(password);
                                user.Save();
                                string logEntry = $"User \"{user}\" Password has been changed sucessfully at {DateTime.Now}\n";
                                Console.WriteLine(logEntry);
                                auditLogManager.Log(logEntry);
                                user.Dispose();
                            }// end of if statement
                            else
                            {
                                Console.WriteLine("Returning to menu...");
                            }// end of else statement
                        }// end of if statement
                        else
                        {
                            Console.WriteLine("Password does not meet the requirement. Please try again.");
                        }
                    }// end of if-statemnet
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine($"User '{username}' not found in Active Directory.");
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }// end of else-statemnet
                }// end of using

            }// end of try
            catch(Exception ex)
            {
                Console.WriteLine($"Error resetting password: {ex.Message}");
            }// end of catch
        }// end of ReserUserPassword

        /// <summary>
        /// A method that vailidate the password meet the requirement of: 
        ///            15 Characters, Symbols, Number, Lower and upper case.
        /// </summary>
        /// <param name="password"></param>
        /// <returns></returns>
        private bool IsPasswordVaild(string password)
        {
            if(password.Length < 15)
            {
                Console.WriteLine($"\nThe password \"{password}\" is less than 15 characters!\n");
                return false;
            }
            else
            {
                bool hasUpperCase = Regex.IsMatch(password, "[A-Z]");
                bool hasLowerCase = Regex.IsMatch(password, "[a-z]");
                bool hasDigit = Regex.IsMatch(password, "[0-9]");
                bool hasSymbol = Regex.IsMatch(password, @"[\W_]");

                if (hasLowerCase == false)
                {
                    Console.WriteLine($"\nThe password \"{password}\" does not have lowercase letters!");
                }
                if (hasUpperCase == false)
                {
                    Console.WriteLine($"\nThe password \"{password}\" does not have uppercase letters!");
                }
                if (hasDigit == false)
                {
                    Console.WriteLine($"\nThe password \"{password}\" does not have digits!");
                }
                if(hasSymbol == false)
                {
                    Console.WriteLine($"The password\"{password}\" does not have symbols!");
                }
                return hasUpperCase && hasLowerCase && hasDigit && hasSymbol;
            }
        }
        /// <summary>
        /// A method return user password expiration date and last time it was set. 
        /// </summary>
        public void GetPasswordExpirationDate()
        {
            Console.Write("Enter the username to check password expiration: ");
            string username = Console.ReadLine();

            try
            {
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);                                           // Searching for the user in AD

                    if (user != null)
                    {
                        DateTime expirationDate = GetPasswordExpirationDate(user);                                                                               // Calculate password experation date
                        DateTime lastSetDate = GetPasswordLastSetDate(user);                                                                                     // Calculate password last time it was set

                        Console.ForegroundColor = ConsoleColor.DarkCyan;
                        Console.WriteLine($"\tPassword last set date for user '{username}': {lastSetDate}");
                        Console.ForegroundColor = ConsoleColor.Gray;

                        if (expirationDate != DateTime.MinValue && user.PasswordNeverExpires == false)
                        {
                            Console.ForegroundColor = ConsoleColor.DarkCyan;
                            Console.WriteLine($"\tPassword expiration date for user '{username}': {expirationDate}");
                            Console.ForegroundColor = ConsoleColor.Gray;
                        }// end inner if-statement
                        if (user.PasswordNeverExpires)
                        {
                            Console.ForegroundColor = ConsoleColor.DarkYellow;
                            Console.WriteLine($"Password for user '{username}' never expires.");
                            Console.ForegroundColor = ConsoleColor.Gray;
                        }// end of inner if-statement
                        
                    }// end of outter if-satetment
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine($"User '{username}' not found in Active Directory.");
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }// end of else-statement
                }// end of using
            }// end of Try-Catch
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: {ex.Message}", ConsoleColor.Red);
                Console.ForegroundColor = ConsoleColor.Gray;
            }// end of catch
        }// end of GetPasswordExpirationDate

        /// <summary>
        /// A method return password exipration date object for a user
        /// </summary>
        /// <param name="user"> Uses user Object in AD</param>
        /// <returns> Password expiration date</returns>
        public DateTime GetPasswordExpirationDate(UserPrincipal user)
        {
            DirectoryEntry deUser = (DirectoryEntry)user.GetUnderlyingObject();                                                               // Grab the underlyning object for the user from AD.
            ActiveDs.IADsUser nativeDeUser = (ActiveDs.IADsUser)deUser.NativeObject;                                                          // Get the native object from AD for the user.

            DateTime passwordExpirationDate = nativeDeUser.PasswordExpirationDate;                                                            // Get password expiration date

            return passwordExpirationDate;                                                                                                    // return the password expiration date for the user.
        }// end of CalcPasswordExpirationDate

        /// <summary>
        /// A method return password last time it was set object for a user
        /// </summary>
        /// <param name="user"> Uses user Object in AD</param>
        /// <returns>Password last changed</returns>
        public DateTime GetPasswordLastSetDate(UserPrincipal user)
        {
            DirectoryEntry deUser = (DirectoryEntry)user.GetUnderlyingObject();
            ActiveDs.IADsUser nativeDeUser = (ActiveDs.IADsUser)deUser.NativeObject;
            DateTime passwordLastChanged = nativeDeUser.PasswordLastChanged;
            
            return passwordLastChanged;                                                                                                       // Return the password last time it changed for the user
        }// end of getPaswordLastSetDate

        /// <summary>
        /// A method to hide every key press for password input
        /// </summary>
        /// <returns>Password input</returns>
        public static string GetPassword()                                                                                                    // Method to read password without displaying it on the console
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
}// end of namespace
