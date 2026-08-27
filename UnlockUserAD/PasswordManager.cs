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
            AppLog.Prompt("Enter the username to reset password for: ");
            string username = ConsoleInput.ReadTrimmed();

            try
            {
                // Bound as the admin account, not the interactive user -- otherwise the audit log
                // names the admin while AD records whoever is logged into the workstation.
                using (PrincipalContext context = AdminSession.CreateContext())
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);                                           // Searching for the user in AD

                    if (user != null)
                    {
                        AppLog.Screen($"\nPassword Requirement: {MinimumPasswordLength} Characters, Symbols, Number, Lower and upper case. ");
                        AppLog.Prompt("Enter the desired password: ");
                        // Masked, and never echoed back. This used to read with Console.ReadLine()
                        // and then print the password to the screen for confirmation.
                        string password = GetPassword();

                        if (IsPasswordVaild(password))
                        {
                            AppLog.Prompt($"Password accepted ({password.Length} characters).\n" +
                                          $"Set this password for '{username}'?(Y/N)");
                            string comfirmation = ConsoleInput.ReadTrimmedUpper();

                            if (comfirmation == "Y")
                            {

                                user.SetPassword(password);
                                user.Save();

                                // A reset password is spoken aloud or typed into a ticket, so it must
                                // not stay valid. The new-hire path already did this; this one did
                                // not, so a helpdesk reset left the temporary password good for the
                                // full domain max password age.
                                bool mustChange = false;
                                try
                                {
                                    user.ExpirePasswordNow();
                                    mustChange = true;
                                }
                                catch (Exception ex)
                                {
                                    ConsoleUi.Warn($"Could not flag the password for change at next logon: {ex.Message}");
                                    ConsoleUi.Note("Tick 'User must change password at next logon' manually in ADUC.");
                                }

                                string logEntry = mustChange
                                    ? $"Password for \"{user}\" reset; must be changed at next logon"
                                    : $"Password for \"{user}\" reset; *** could NOT be flagged for change at next logon ***";
                                ConsoleUi.Ok(logEntry);
                                auditLogManager?.Log(logEntry); // null-safe: auditLogManager may be null when using default constructor
                                user.Dispose();
                            }// end of if statement
                            else
                            {
                                AppLog.Screen("Returning to menu...");
                            }// end of else statement
                        }// end of if statement
                        else
                        {
                            AppLog.Warn("Password does not meet the requirement. Please try again.");
                        }
                    }// end of if-statemnet
                    else
                    {
                        AppLog.Warn($"User '{username}' not found in Active Directory.", color: Color.IndianRed);
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
                AppLog.Warn("\nNo password was entered!\n", color: Color.IndianRed);
                return false;
            }
            if (password.Length < MinimumPasswordLength)
            {
                AppLog.Warn($"\nThe password is {password.Length} characters — less than the required {MinimumPasswordLength}!\n", color: Color.IndianRed);
                return false;
            }

            bool hasUpperCase = Regex.IsMatch(password, "[A-Z]");
            bool hasLowerCase = Regex.IsMatch(password, "[a-z]");
            bool hasDigit = Regex.IsMatch(password, "[0-9]");
            bool hasSymbol = Regex.IsMatch(password, @"[\W_]");

            if (!hasLowerCase)
            {
                AppLog.Warn("\nThe password does not have lowercase letters!", color: Color.IndianRed);
            }
            if (!hasUpperCase)
            {
                AppLog.Warn("\nThe password does not have uppercase letters!", color: Color.IndianRed);
            }
            if (!hasDigit)
            {
                AppLog.Warn("\nThe password does not have digits!", color: Color.IndianRed);
            }
            if (!hasSymbol)
            {
                AppLog.Warn("\nThe password does not have symbols!", color: Color.IndianRed);
            }
            return hasUpperCase && hasLowerCase && hasDigit && hasSymbol;
        }
        /// <summary>
        /// A method return user password expiration date and last time it was set. 
        /// </summary>
        public void GetPasswordExpirationDate()
        {
            AppLog.Prompt("Enter the username to check password expiration: ");
            string username = ConsoleInput.ReadTrimmed();

            try
            {
                // Bound as the admin account, not the interactive user -- otherwise the audit log
                // names the admin while AD records whoever is logged into the workstation.
                using (PrincipalContext context = AdminSession.CreateContext())
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);                                           // Searching for the user in AD

                    if (user != null)
                    {
                        DateTime expirationDate = GetPasswordExpirationDate(user);                                                                               // Calculate password experation date
                        DateTime? lastSetDate = GetPasswordLastSetDate(user);                                                                                     // Calculate password last time it was set

                        AppLog.Screen($"\tPassword last set date for user '{username}': {lastSetDate}", Color.DarkCyan);

                        if (expirationDate != DateTime.MinValue && user.PasswordNeverExpires == false)
                        {
                            AppLog.Screen($"\tPassword expiration date for user '{username}': {expirationDate}", Color.DarkCyan);
                        }// end inner if-statement
                        if (user.PasswordNeverExpires)
                        {
                            AppLog.Screen($"Password for user '{username}' never expires.", Color.DarkGoldenrod);
                        }// end of inner if-statement

                    }// end of outter if-satetment
                    else
                    {
                        AppLog.Warn($"User '{username}' not found in Active Directory.", color: Color.IndianRed);
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
        /// The attribute AD computes per user for when their password expires. Constructed, so it
        /// is not returned unless explicitly requested -- see the RefreshCache call below.
        /// </summary>
        private const string PasswordExpiryAttribute = "msDS-UserPasswordExpiryTimeComputed";

        /// <summary>
        /// When a password expires, as Active Directory itself computes it.
        ///
        /// This used to be calculated as pwdLastSet + the *domain* maxPwdAge, which ignores
        /// fine-grained password policies. This domain has one -- Lloyd_Password_Policy, 90 days,
        /// applied to Domain Users -- overriding a 30-day domain default, so every expiry was
        /// reported 60 days early. Asking AD for the computed value honours any PSO automatically
        /// and needs no policy enumeration.
        /// </summary>
        /// <returns>The expiry date, or DateTime.MinValue when the password does not expire.</returns>
        public DateTime GetPasswordExpirationDate(UserPrincipal user)
        {
            if (user == null) return DateTime.MinValue;
            if (user.PasswordNeverExpires) return DateTime.MinValue;

            try
            {
                DirectoryEntry deUser = (DirectoryEntry)user.GetUnderlyingObject();

                // Constructed attributes are absent from the default property cache; without this
                // the value comes back empty and every account looks like it never expires.
                deUser.RefreshCache(new[] { PasswordExpiryAttribute });

                long? fileTime = ConvertLargeIntegerToLong(deUser.Properties[PasswordExpiryAttribute].Value);
                if (fileTime.HasValue)
                {
                    // Two sentinels, neither of which is a real date: long.MaxValue means the
                    // password never expires, 0 means it must be changed at next logon. Passing
                    // either to FromFileTimeUtc throws.
                    if (fileTime.Value == long.MaxValue || fileTime.Value <= 0) return DateTime.MinValue;

                    return DateTime.FromFileTimeUtc(fileTime.Value).ToLocalTime();
                }

                // The attribute is unavailable (pre-2008 functional level, or no read access).
                // Fall back to the old arithmetic, which is PSO-blind and so may read early.
                AppLog.Detail($"{PasswordExpiryAttribute} unavailable for '{user.SamAccountName}'; falling back to the domain maximum password age.");

                DateTime? pwdLastSet = ConvertLargeIntegerToDateTime(deUser.Properties["pwdLastSet"].Value);
                TimeSpan? maxPwdAge = GetDomainMaxPasswordAge(user.Context);
                if (!pwdLastSet.HasValue || !maxPwdAge.HasValue) return DateTime.MinValue;

                return pwdLastSet.Value.AddTicks(Math.Abs(maxPwdAge.Value.Ticks));
            }
            catch (Exception ex)
            {
                // Report rather than swallow: a rights failure reading the attribute used to be
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

        /// <summary>
        /// The domain-wide maximum password age.
        ///
        /// FALLBACK ONLY -- do not make this the primary source again. It reads the domain object,
        /// which is blind to fine-grained password policies: this domain's PSO sets 90 days while
        /// the domain object says 30, so relying on this reported every expiry 60 days early.
        /// <see cref="GetPasswordExpirationDate"/> asks AD for the per-user computed value instead.
        /// </summary>
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
