# Active Directory Utility

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
4. Enter the admin username and password when prompted.
5. Follow the instructions:

    ```
    Type "1" and hit Enter to unlock specific user. 
    Type "2" and hit Enter to check if any accounts are locked.
    Type "3" and hit Enter to unlock all locked accounts.
    Type "4" and hit Enter to list all groups in active directory.
    Type "5" and hit Enter to add user to a group.
    Type "6" and hit Enter to check the password expiration date.
    Type "7" and hit enter to exit/close the application
    ```

## Features

- **Unlock Specific User:** Allows unlocking a specific user account by entering the username.
- **Check Locked Accounts:** Displays all locked user accounts in Active Directory.
- **Unlock All Locked Accounts:** Unlocks all locked user accounts in Active Directory.
- **List All Groups:** Lists all groups in the Active Directory.
- **Add User to Group:** Adds a user to a specified group in the Active Directory.
- **Password Expiration:** Display a specific user password expiration date. 

## Important Notes

- Ensure that you have the necessary permissions to unlock user accounts and manage groups in Active Directory.
- This program does not store or log any credentials provided.
