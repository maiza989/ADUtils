using Pastel;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Drawing;


namespace ADUtils
{
    public class AccountDeactivationManager
    {
        AccountCreationManager ACManager;

        // TODO - Fix accountcreationmanager null reference exception
        /*public AccountDeactivationManager()
        {
            this.ACManager = new AccountCreationManager(Program.configuration);
        }
*/
        public void DeactivateUserAccount(PrincipalContext context, string adminUsername, string adminPassword)
        {
            
            string ouPath = $"LDAP://OU=Ex Employee,OU={ACManager._myCompany}_Lou,DC={ACManager._myDomain},DC={ACManager._myDomain}";
            DateTime deletionDate = DateTime.Now.AddDays(31);                                                                           // Calculate Today's date + 31 days
            string deletionDateString = deletionDate.ToString("MM-dd-yyyy");                                                            // Format the date
            bool returnToMenu = false;

            do
            {
                Console.Write($"Enter the username to deactivate (type {"'exit'".Pastel(Color.MediumPurple)} to return to the main menu): ");
                string username = Console.ReadLine().Trim().ToLower();

                if (username.ToLower().Trim() == "exit")
                {
                    returnToMenu = true;
                }// end of if statement
                else
                {
                    try
                    {
                        UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);              // Search for specific user using username
                        if(user != null)
                        {
                            var groups = user.GetGroups();                                                                              // Get all group user is member of
                            foreach (var group in groups )
                            {
                                if(group.Name != "Domain Users" && group is GroupPrincipal)                                             // Remove all group except 'Domain User'
                                {
                                    GroupPrincipal groupPrincipal = (GroupPrincipal)group;
                                    groupPrincipal.Members.Remove(user);
                                    groupPrincipal.Save();
                                }// end of if statement
                            }// end of foreach
                            Console.WriteLine($"User account '{username}' has been removed from all groups except 'Domain Users'".Pastel(Color.LimeGreen));
                            if (user.Enabled == true)
                            {
                                user.Enabled = false;                                                                                   // Disabling the user account
                                user.Description = $"Delete on {deletionDateString}";                                                   // Change description with reminder of when to delete the ex user account
                                user.Save();
                                Console.WriteLine($"User account '{username}' has been disabled\nAccount description changed to 'Delete on {deletionDateString}!".Pastel(Color.LimeGreen));

                                using (DirectoryEntry userEntry = (DirectoryEntry)user.GetUnderlyingObject())                           // Move user to the specified OU
                                {
                                    DirectoryEntry startOU = new DirectoryEntry(userEntry.Path);
                                    DirectoryEntry endOU = new DirectoryEntry(ouPath, adminUsername, adminPassword);
                                    
                                    userEntry.CommitChanges();
                                    startOU.MoveTo(endOU);
                                    Console.WriteLine($"User account '{username}' has been moved to Ex Employee OU".Pastel(Color.LimeGreen));
                                }// end of using
                            }// end of if statement
                            else
                            {
                                Console.WriteLine($"User account '{username}' is {"ALREADY".Pastel(Color.MediumPurple)} disabled".Pastel(Color.DarkGoldenrod));
                            }// end of else statement
                        }// end of if statement
                        else
                        {
                            Console.WriteLine($"\tUser account '{username}' not found in Active Directory.".Pastel(Color.IndianRed));
                        }// end of else statement
                    }// end of try
                    catch
                    {

                    }// end of catch
                }// end of else statement
            } while(!returnToMenu);   
        }// end of DeactivateUserAccount
    }// end of class
}// end of namespace
