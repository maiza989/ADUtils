using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Text.RegularExpressions;
using System.Reflection;
using Pastel;
using System.Drawing;

namespace ADUtils
{
    public class PasswordManager
    {
        /// <summary>
        /// Minimum password length enforced by <see cref="IsPasswordVaild"/>. Referenced by the
        /// prompt text too, so the stated requirement and the check can't drift apart.
        /// </summary>
        internal const int MinimumPasswordLength = 15;

        AuditLogManager auditLogManager;

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
            string username = ConsoleInput.ReadTrimmed();

            try
            {
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);                                           // Searching for the user in AD

                    if (user != null)
                    {
                        Console.WriteLine($"\nPassword Requirement: {MinimumPasswordLength} Characters, Symbols, Number, Lower and upper case. ");
                        Console.Write("Enter the desired password: ");
                        // Masked, and never echoed back. This used to read with Console.ReadLine()
                        // and then print the password to the screen for confirmation.
                        string password = GetPassword();

                        if (IsPasswordVaild(password))
                        {
                            Console.Write($"Password accepted ({password.Length} characters).\n" +
                                              $"Set this password for '{username}'?(Y/N)");
                            string comfirmation = ConsoleInput.ReadTrimmedUpper();

                            if (comfirmation == "Y")
                            {

                                user.SetPassword(password);
                                user.Save();
                                string logEntry = $"User \"{user}\" Password has been changed successfully at {DateTime.Now}\n";
                                Console.WriteLine(logEntry);
                                auditLogManager?.Log(logEntry); // null-safe: auditLogManager may be null when using default constructor
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
            catch (Exception ex)
            {
                AppLog.Error($"Error resetting password: {ex.Message}", ex);
            }// end of catch
        }// end of ReserUserPassword

        /// <summary>
        /// A method that vailidate the password meet the requirement of: 
        ///            15 Characters, Symbols, Number, Lower and upper case.
        /// </summary>
        /// <param name="password"></param>
        /// <returns></returns>
        internal static bool IsPasswordVaild(string password)
        {
            // Report which rule failed, never the attempted password. Echoing it here meant a
            // rejected password appeared on screen up to five times, and would have been written
            // to the audit log verbatim if console redirection were ever switched on.
            if (string.IsNullOrEmpty(password))
            {
                Console.WriteLine("\nNo password was entered!\n".Pastel(Color.IndianRed));
                return false;
            }
            if (password.Length < MinimumPasswordLength)
            {
                Console.WriteLine($"\nThe password is {password.Length} characters — less than the required {MinimumPasswordLength}!\n".Pastel(Color.IndianRed));
                return false;
            }

            bool hasUpperCase = Regex.IsMatch(password, "[A-Z]");
            bool hasLowerCase = Regex.IsMatch(password, "[a-z]");
            bool hasDigit = Regex.IsMatch(password, "[0-9]");
            bool hasSymbol = Regex.IsMatch(password, @"[\W_]");

            if (!hasLowerCase)
            {
                Console.WriteLine("\nThe password does not have lowercase letters!".Pastel(Color.IndianRed));
            }
            if (!hasUpperCase)
            {
                Console.WriteLine("\nThe password does not have uppercase letters!".Pastel(Color.IndianRed));
            }
            if (!hasDigit)
            {
                Console.WriteLine("\nThe password does not have digits!".Pastel(Color.IndianRed));
            }
            if (!hasSymbol)
            {
                Console.WriteLine("\nThe password does not have symbols!".Pastel(Color.IndianRed));
            }
            return hasUpperCase && hasLowerCase && hasDigit && hasSymbol;
        }
        /// <summary>
        /// A method return user password expiration date and last time it was set. 
        /// </summary>
        public void GetPasswordExpirationDate()
        {
            Console.Write("Enter the username to check password expiration: ");
            string username = ConsoleInput.ReadTrimmed();

            try
            {
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain))
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);                                           // Searching for the user in AD

                    if (user != null)
                    {
                        DateTime expirationDate = GetPasswordExpirationDate(user);                                                                               // Calculate password experation date
                        DateTime? lastSetDate = GetPasswordLastSetDate(user);                                                                                     // Calculate password last time it was set

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
                // The stray ConsoleColor argument here selected WriteLine(string format, object
                // arg0), so the already-interpolated message was re-parsed as a format string --
                // any LDAP/COM message containing braces threw FormatException from the handler.
                AppLog.Error($"Error: {ex.Message}", ex, Color.Crimson);
            }// end of catch
        }// end of GetPasswordExpirationDate

        /// <summary>
        /// A method return password exipration date object for a user
        /// </summary>
        /// <param name="user"> Uses user Object in AD</param>
        /// <returns> Password expiration date</returns>
        public DateTime GetPasswordExpirationDate(UserPrincipal user)
        {
            if (user == null) return DateTime.MinValue;

            try
            {
                DirectoryEntry deUser = (DirectoryEntry)user.GetUnderlyingObject();
                DateTime? pwdLastSet = ConvertLargeIntegerToDateTime(deUser.Properties["pwdLastSet"].Value);

                if (!pwdLastSet.HasValue || user.PasswordNeverExpires)
                    return DateTime.MinValue;

                TimeSpan? maxPwdAge = GetDomainMaxPasswordAge(user.Context);
                if (!maxPwdAge.HasValue)
                    return DateTime.MinValue;

                // maxPwdAge is stored as a negative timespan; add absolute value to last set
                DateTime expiration = pwdLastSet.Value.AddTicks(Math.Abs(maxPwdAge.Value.Ticks));
                return expiration;
            }
            catch (Exception ex)
            {
                // Report rather than swallow: a rights failure reading pwdLastSet used to be
                // indistinguishable from "the password never expires".
                AppLog.Warn($"Could not determine password expiration for '{user.SamAccountName}': {ex.Message}", ex, Color.DarkGoldenrod);
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// A method return password last time it was set object for a user
        /// </summary>
        /// <param name="user"> Uses user Object in AD</param>
        /// <returns>Password last changed</returns>
        public DateTime? GetPasswordLastSetDate(UserPrincipal user)
        {
            if (user == null) return null;

            try
            {
                DirectoryEntry deUser = (DirectoryEntry)user.GetUnderlyingObject();
                return ConvertLargeIntegerToDateTime(deUser.Properties["pwdLastSet"].Value);
            }
            catch (Exception ex)
            {
                AppLog.Warn($"Could not read pwdLastSet for '{user.SamAccountName}': {ex.Message}", ex, Color.DarkGoldenrod);
                return null;
            }
        }

        private DateTime? ConvertLargeIntegerToDateTime(object largeInt)
        {
            if (largeInt == null) return null;

            try
            {
                // The COM LargeInteger object exposes HighPart and LowPart properties
                var type = largeInt.GetType();
                var high = (int)type.InvokeMember("HighPart", BindingFlags.GetProperty, null, largeInt, null);
                var low  = (int)type.InvokeMember("LowPart",  BindingFlags.GetProperty, null, largeInt, null);

                long fileTime = ((long)high << 32) + (uint)low;
                if (fileTime <= 0) return null;
                return DateTime.FromFileTimeUtc(fileTime).ToLocalTime();
            }
            catch
            {
                return null;
            }
        }

        private long? ConvertLargeIntegerToLong(object largeInt)
        {
            if (largeInt == null) return null;
            try
            {
                var type = largeInt.GetType();
                var high = (int)type.InvokeMember("HighPart", BindingFlags.GetProperty, null, largeInt, null);
                var low  = (int)type.InvokeMember("LowPart",  BindingFlags.GetProperty, null, largeInt, null);
                long value = ((long)high << 32) + (uint)low;
                return value;
            }
            catch
            {
                return null;
            }
        }

        private TimeSpan? GetDomainMaxPasswordAge(PrincipalContext context)
        {
            try
            {
                // Read defaultNamingContext from RootDSE, then domain object maxPwdAge
                using (DirectoryEntry rootDse = new DirectoryEntry("LDAP://RootDSE"))
                {
                    string defaultNamingContext = rootDse.Properties["defaultNamingContext"].Value as string;
                    if (string.IsNullOrEmpty(defaultNamingContext)) return null;

                    using (DirectoryEntry domain = new DirectoryEntry($"LDAP://{defaultNamingContext}"))
                    {
                        object maxPwdAgeObj = domain.Properties["maxPwdAge"].Value;
                        long? ticks = ConvertLargeIntegerToLong(maxPwdAgeObj);
                        if (!ticks.HasValue) return null;

                        // maxPwdAge is negative; create a TimeSpan from absolute ticks
                        return TimeSpan.FromTicks(Math.Abs(ticks.Value));
                    }
                }
            }
            catch (Exception ex)
            {
                AppLog.Warn($"Could not read the domain maximum password age: {ex.Message}", ex, Color.DarkGoldenrod);
                return null;
            }
        }

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
