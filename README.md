# Active Directory Utility

 =>>**PROJECT IS IN DEVELOPMENT**<<=

This C# program allows an admin to perform actions on users in Active Directory using admin credentials. It prompts for the admin username and password, then provides an interface to unlock user accounts, check locked accounts, list all security groups and distribution lists, add a user to security and distribution lists, and check user password expiration date.

## Prerequisites

- .NET Framework installed.
- System.DirectoryServices.AccountManagement.
- ActiveDs library.
- Access to an Active Directory environment.

## How to Use

1. Clone or download this repository.
2. Compile the program using Visual Studio or any C# compiler.
3. Run the compiled executable.
4. Enter your Active Directory admin username and password when prompted.
5. Follow the instructions:

    ```
    Type "1" and hit Enter to access user lock out manager 
    Type "2" and hit Enter to access user group manager.
    Type "3" and hit Enter to access user information manager.
    ```
## Features

#### Lock Out Manager
- **Unlock Specific User:** Allows unlocking a specific user account by entering the username.
- **Check Locked Accounts:** Displays all locked user accounts in Active Directory.
- **Unlock All Locked Accounts:** Unlocks all locked user accounts in Active Directory.
#### Group Manager
- **List All Groups:** Lists all groups in the Active Directory.
- **Add User to Group:** Adds a user to a specified group in the Active Directory.
- **Remove User From group:** Remove a user from a specific group in the Active Directory
- **Check Group Members:** Check who is a memeber of a specific group in the Active Directory
- **Email Notification:** Send an email notifcation to desired location when update group member in Active Directory.
#### User Info Manager
- **Password Expiration:** Display a specific user password expiration date.
- **User Infromation**: Display gerenal user infromamtion from the Active Direcotry.

## Important Notes

- Ensure that you have the necessary permissions to unlock user accounts and manage groups in Active Directory.
- This program does not store or log any credentials provided.
