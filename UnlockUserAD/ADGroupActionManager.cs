using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;


namespace UnlockUserAD
{
    public class ADGroupActionManager
    {
       

        /// <summary>
        /// A method to add user to group security and distrbuiton list in Active Directory.
        /// </summary>
        /// <param name="context"></param>
        public void AddUserToGroup(PrincipalContext context)
        {
            bool isExit = false;
            do
            {
                Console.Write("Enter the username(Type 'exit' to go back to menu): ");
                string username = Console.ReadLine().Trim();
                if (username.ToLower().Trim() == "exit")
                {
                    isExit = true;
                    Console.WriteLine($"\nReturing to menu...");
                    break;
                }
                Console.Write("Enter the group name (Type 'exit' to go back to menu): ");
                string groupName = Console.ReadLine().Trim();
                if (groupName.ToLower().Trim() == "exit")
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
            } while (!isExit);
        }// end of AddUserToGroup

        /// <summary>
        /// A method that remove a user from a security group and distrubtion list in Active Directory
        /// </summary>
        /// <param name="context"></param>
        public void RemoveUserToGroup(PrincipalContext context)
        {
            bool isExit = false;
            do
            {
                Console.Write("Enter the username(Type 'exit' to go back to menu): ");
                string username = Console.ReadLine().Trim();
                if (username.ToLower().Trim() == "exit")
                {
                    isExit = true;
                    Console.WriteLine($"\nReturing to menu...");
                    break;
                }
                Console.Write("Enter the group name (Type 'exit' to go back to menu): ");
                string groupName = Console.ReadLine().Trim();
                if (groupName.ToLower().Trim() == "exit")
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
                                if (group.Members.Contains(user))                                                                                                      // If the user is not in the group add him
                                {
                                    // Add the user to the group
                                    group.Members.Remove(user);                                                                                                      // Apply changes
                                    group.Save();
                                    group.Dispose();
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine($"User '{username}' removed from group '{groupName}' successfully.");
                                    Console.ForegroundColor = ConsoleColor.Gray;
                                }// end of inner-2 if-statement
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                                    Console.WriteLine($"User '{username}' is not a member of group '{groupName}'.");
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
                        Console.WriteLine($"Error removing user from group: {ex.Message}");
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }// end of catch
                }
            } while (!isExit);
        }// end of AddUserToGroup

        /// <summary>
        /// A method that list all group secuirty and distrubtion list in Active Directory.
        /// </summary>
        /// <param name="context"></param>
        public void ListAllGroups(PrincipalContext context)
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

        /// <summary>
        /// A method that searches for the members of a specified group in Active Directory and lists them in a grid style.
        /// </summary>
        /// <param name="context"></param>
        public void ListGroupMembers(PrincipalContext context)
        {
            bool isExit = false;
            do
            {
                Console.Write("Enter the group name (Type 'exit' to go back to menu): ");
                string groupName = Console.ReadLine().Trim();
                if (groupName.ToLower() == "exit")
                {
                    isExit = true;
                    Console.WriteLine("\nReturning to menu...");
                    break;
                }
                else
                {
                    try
                    {
                        GroupPrincipal group = GroupPrincipal.FindByIdentity(context, groupName); // Check for group in AD

                        if (group != null)
                        {
                            Console.WriteLine($"\nMembers of group '{groupName}':");

                            List<string> memberNames = new List<string>();
                            foreach (var member in group.GetMembers())
                            {
                                memberNames.Add(member.SamAccountName);
                            }// end of foreach

                            if (memberNames.Count > 0)
                            {
                                memberNames.Sort();
                                int maxMemberNameLength = memberNames.Max(m => m.Length);                                                                   // Find the longest member name length
                                int columnWidth = maxMemberNameLength + 5;                                                                                  // Add padding
                                int numColumns = Console.WindowWidth / columnWidth;                                                                         // Calculate number of columns based on window width
                                int numRows = (int)Math.Ceiling((double)memberNames.Count / numColumns);                                                    // Calculate number of rows

                                for (int i = 0; i < numRows; i++)                                                                                           // Nested for loop to print member names in a grid style
                                {
                                    for (int j = 0; j < numColumns; j++)
                                    {
                                        int index = i + j * numRows;                                                                                        // Calculate the index of the member based on 'i' rows and 'j' columns
                                        if (index < memberNames.Count)
                                        {
                                            Console.Write($"- {memberNames[index].PadRight(columnWidth)}");                                                 // Print each member name with specified right padding
                                        }// ned of if statement
                                    }// end of for loop
                                    Console.WriteLine();
                                }// end of of for loop 
                            }// end of if-statement
                            else
                            {
                                Console.WriteLine("No members found in this group.");
                            }// end of else-statement
                        }// end of if-statements
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.DarkRed;
                            Console.WriteLine($"Group '{groupName}' not found in Active Directory.");
                            Console.ForegroundColor = ConsoleColor.Gray;
                        }// end of else
                    }// end of try-catch
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Error listing members of group: {ex.Message}");
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }// end of catch
                }// end of else
            } while (!isExit);
        }// end of ListGroupMembers
    }// end of class
}// end of namespace
