using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using ADUtils;
using Pastel;
using System.Drawing;
using Microsoft.Extensions.Configuration;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography;
// TODO - DONE - Audit Logging: Log fuction to record important action performed.
// TODO - DONE User Account Deactivation: Implement functionality to deactivate user accounts securely.
// TODO - Create/Delete Groups: Allow creating and deleting security groups or distribution lists.


/// <summary>
/// Edit the following:
/// 
///  - BaseLogDirectory with your desire Log location in AuditLogManager Class.
/// 
///  - myDomain, myDomainDotCom, myParentOU, myCompany, myExchangeDatabase, myExchangeServer, and outPath with your own values for successful execution in AccountCreationManager Class.
/// 
///  - mySTMPServer, myFromEmail, myPassword, myToEmail with your own values for successful execution in EmailNotificationManager Class. 
///  
/// </summary>
class Program
{

    static bool isLocked = false;
    static int countdownSeconds = 60;
    public static string adminUsername;
    private static string adminPassword;
    static private bool isAuthenticated = false;
    public static IConfiguration configuration;

   /* static void GetAdminCreditials()
    {

        Console.Write("Enter admin username: ");
        adminUsername = Console.ReadLine().Trim();
        Console.Write("Enter admin password: ");
        adminPassword = PasswordManager.GetPassword().Trim();
    }*/

    static X509Certificate2 GetAdminCertificate()
    {
        Console.Write("Enter admin username: ");
        adminUsername = Console.ReadLine().Trim();

       /* Console.Write("Enter your smart card PIN: ");
        string smartCardPin = PasswordManager.GetPassword().Trim();*/

        using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
        {
            store.Open(OpenFlags.ReadOnly);

           /* // Find certificates that are valid, have the correct key usage, and potentially filter by Issuer or Subject Name
            X509Certificate2Collection certCollection = store.Certificates                                                  
                .Find(X509FindType.FindByTimeValid, DateTime.Now, false)
                .Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, false);
           */

            // Optionally add more filtering criteria specific to your YubiKey
            X509Certificate2Collection yubiKeyCerts = new X509Certificate2Collection();
            foreach (var cert in store.Certificates)
            {
                // First, check if the subject name matches the admin username
                string subjectName = cert.GetNameInfo(X509NameType.SimpleName, false);
                if (!subjectName.Equals(adminUsername, StringComparison.OrdinalIgnoreCase))
                {
                    // Skip this certificate if the subject does not match
                    continue;
                }
                try
                {
                    if (cert.HasPrivateKey)
                    {
                        // Access the private key
                        var privateKey = cert.PrivateKey;

                        if (privateKey is RSACng rsaCng)
                        {
                            if (rsaCng.KeyExchangeAlgorithm.Equals("RSA", StringComparison.OrdinalIgnoreCase))
                            {
                                // Attempt to get the Key Storage Provider (KSP) information, if it exists
                                if (subjectName.Equals(adminUsername, StringComparison.OrdinalIgnoreCase))
                                {
                                    // Attempt to sign data with the RSA key to validate access
                                    byte[] dataToSign = new byte[] { 0x01 }; // Dummy data
                                    byte[] signedData = rsaCng.SignData(dataToSign, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                                    yubiKeyCerts.Add(cert);
                                }// end of if-statement
                            }
                        }
                    }// end of if-statement
                }// end of try
                catch
                {
                    continue;
                }// end of catch
            }// end of foreach

            if (yubiKeyCerts.Count == 0)
            {
                Console.WriteLine("No smart card detected or no valid certificates found on the connected smart card.");
                return null;
            }

            X509Certificate2 selectedCert = null;
            if (yubiKeyCerts.Count > 0)
            {
                selectedCert = X509Certificate2UI.SelectFromCollection(
                    yubiKeyCerts,
                    "Select a YubiKey certificate",
                    "Please select your admin certificate from the YubiKey",
                    X509SelectionFlag.SingleSelection
                )[0];
            }
            adminUsername = selectedCert.GetNameInfo(X509NameType.SimpleName, false);
            adminPassword = selectedCert.PrivateKey.SignatureAlgorithm;
            return selectedCert;
        }// end of using x509Store
    }// end of GetAdminCertificate

    static void Main(string[] args)
    {
        ActiveDirectoryManager ADManager = new ActiveDirectoryManager();
        AccountCreationManager ACManager;
        AccountDeactivationManager ACCDeactivationManager = new AccountDeactivationManager();
        PasswordManager PWDManager = null;
        ADGroupActionManager ADGroupManager = null;
        AuditLogManager auditLogManager = null;

        configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("Appsettings.json", optional: false, reloadOnChange: true)
            .Build();
        EmailNotifcationManager emailManager = new EmailNotifcationManager(configuration);
        string _myDomainName = configuration["AccountCreationSettings:myDomainName"];

        do
        {
        GetAdminCertificate();
            
            try
            {
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain, _myDomainName))                                               // Check if the the password/user are correct
                {
                    if (context.ConnectedServer != null)                                                                                                                      // Throw error if the password/username is incorrect        
                    {
                        isAuthenticated = true;
                        Console.WriteLine($"Connected to Active Directory as: {adminUsername}.".Pastel(Color.GreenYellow));
               
                        auditLogManager = new AuditLogManager(adminUsername, configuration);
                        ADGroupManager = new ADGroupActionManager(auditLogManager);
                        PWDManager = new PasswordManager(auditLogManager);
                        ACManager = new AccountCreationManager(auditLogManager, configuration);

                        bool exit = false;
                        while (!exit)                                                                                                                                          // Loop the menu
                        {
                            DisplayMainMenu();
                            string choice = Console.ReadLine();
                            exit = HandleMainMenuChoice(choice, context, ADManager, ADGroupManager, PWDManager, ACManager, ACCDeactivationManager);
                        }// end of while-loop
                    }// end of if statement
                    context.Dispose();
                }// end of using
            }// end of Try-Catch
            catch (DirectoryServicesCOMException)                                                                                                                              // Error out if password/username are incorrect
            {         
                Console.WriteLine("Error: Unable to connect to the Active Directory server. Please check your credentials and try again.".Pastel(Color.IndianRed));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}".Pastel(Color.IndianRed));
            }// end of Catch
        } while (!isAuthenticated || string.IsNullOrEmpty(adminUsername));                                                                                                                                            // Repeat until a valid password is entered
    }// end of Main Method


    //--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    //                                                                                                          UI
    //--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    static void DisplayMainMenu()
    {

        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine("\nSelect an option:");
        Console.WriteLine("1. Locked Out Management");
        Console.WriteLine("2. Group Management");
        Console.WriteLine("3. User Information Management");
        Console.WriteLine("4. Exit");
        Console.Write("Enter your choice: ");
    }// end of DisplayMainMenu

    /// <summary>
    /// Main Menu interface
    /// </summary>
    /// <param name="choice"> User choice</param>
    /// <param name="context"> Active directory object</param>
    /// <param name="ADManager"> A class that manage user lockout</param>
    /// <param name="ADGroupManager">A class that manage user groups </param>
    /// <param name="PWDManager"> A class that manager user password related events</param>
    /// <returns></returns>
    static bool HandleMainMenuChoice(string choice, PrincipalContext context, ActiveDirectoryManager ADManager, ADGroupActionManager ADGroupManager, PasswordManager PWDManager, AccountCreationManager ACManager, AccountDeactivationManager ACCDeactivationManager)
    {
        switch (choice)
        {
            case "1":
                DisplayLockedOutMenu(context, ADManager);
                break;
            case "2":
                DisplayGroupManagementMenu(context, ADGroupManager);
                break;
            case "3":
                DisplayUserInfoMenu(context, ADManager, PWDManager, ACManager, ACCDeactivationManager);
                break;
            case "4":
                return true;
            case "clear":
                Console.Clear();
                break;
            default:
                Console.WriteLine("Invalid option. Please try again.".Pastel(Color.IndianRed));
                break;
        }// end of switch case
        return false;
    }// end of Handle

    /// <summary>
    /// A UI that host all user lockout management
    /// </summary>
    /// <param name="context"></param>
    /// <param name="ADManager"></param>
    static void DisplayLockedOutMenu(PrincipalContext context, ActiveDirectoryManager ADManager)
    {
        bool exit = false;
        while (!exit)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("\nLocked Out Manager:");
            Console.WriteLine("1. Unlock a Specific User");
            Console.WriteLine("2. Check All Locked Accounts");
            Console.WriteLine("3. Unlock All Locked Accounts");
            Console.Write($"Enter your choice(Type {"'exit'".Pastel(Color.MediumPurple)} to return to main menu): ");
            string choice = Console.ReadLine().ToLower().Trim();
            switch (choice)
            {
                case "1":
                    ADManager.UnlockUser(context);
                    break;
                case "2":
                    ADManager.CheckLockedAccounts(context);
                    break;
                case "3":
                    ADManager.UnlockAllUsers(context);
                    break;
                case "exit":
                    exit = true;
                    break;
                default:
                    Console.WriteLine("Invalid option. Please try again.".Pastel(Color.IndianRed));
                    break;
            }// end of switch-case
        }// end of while
        Console.Clear();
    }// end of DisplayLockedOutMenu

    /// <summary>
    /// A UI that host all security group and distirbution list management 
    /// </summary>
    /// <param name="context"></param>
    /// <param name="ADGroupManager"></param>
    static void DisplayGroupManagementMenu(PrincipalContext context, ADGroupActionManager ADGroupManager)
    {
        bool exit = false;
        while (!exit)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("\nGroup Management:");
            Console.WriteLine("1. List All Groups in Active Directory");
            Console.WriteLine("2. Add User to a Group");
            Console.WriteLine("3. Remove User From a Group");
            Console.WriteLine("4. Check Who is Member in a Group");
            Console.Write($"Enter your choice(Type {"'exit'".Pastel(Color.MediumPurple)} to return to main menu): ");

            string choice = Console.ReadLine().ToLower().Trim();
            switch (choice)
            {
                case "1":
                    ADGroupManager.ListAllGroups(context);
                    break;
                case "2":
                    ADGroupManager.AddUserToGroup(context);
                    break;
                case "3":
                    ADGroupManager.RemoveUserToGroup(context);
                    break;
                case "4":
                    ADGroupManager.ListGroupMembers(context);
                    break;
                case "exit":
                    exit = true;
                    break;
                default:
                    Console.WriteLine("Invalid option. Please try again.".Pastel(Color.IndianRed));
                    break;
            }// end of switch-case
        }// end of while loop
        Console.Clear();
    }// end of DisplayGroupMangementMenu

    /// <summary>
    /// A UI that host all user info management. 
    /// </summary>
    /// <param name="context"></param>
    /// <param name="ADManager"></param>
    /// <param name="PWDManager"></param>
    static void DisplayUserInfoMenu(PrincipalContext context, ActiveDirectoryManager ADManager, PasswordManager PWDManager, AccountCreationManager ACManager, AccountDeactivationManager ACCDeactivationManager)
    {
        bool exit = false;
        while (!exit)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("\nUser Information:");
            Console.WriteLine("1. Check User Password Expiration Date");
            Console.WriteLine("2. Display General User Info");
            Console.WriteLine("3. Reset A User Password");
            Console.WriteLine("4. Create New User Account");
            Console.WriteLine("5. Disable User Account");
            Console.Write($"Enter your choice(Type {"'exit'".Pastel(Color.MediumPurple)} to return to main menu): ");

            string choice = Console.ReadLine().ToLower().Trim();
            switch (choice)
            {
                case "1":
                    PWDManager.GetPasswordExpirationDate();
                    break;
                case "2":
                    ADManager.DisplayUserInfo(context);
                    break;
                case "3":
                    PWDManager.ResetUserPassowrd();
                    break;
                case "4":
                    ACManager.CreateUserAccount(adminUsername, adminPassword);
                    break;
                case "5":
                   ACCDeactivationManager.DeactivateUserAccount(context, adminUsername, adminPassword);
                    break;
                case "exit":
                    exit = true;
                    break;
                default:
                    Console.WriteLine("Invalid option. Please try again.".Pastel(Color.IndianRed));
                    break;
            }// end of switch-case
        }// end of while loop
        Console.Clear();
    }// end of DisplayUserInfoMenu

    static void DisplayUserCreationMenu()
    {

    }
   
}// end of class
