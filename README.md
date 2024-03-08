# Active Directory User Account Unlocker

This C# program allows you to unlock user accounts in Active Directory using admin credentials. It prompts for the admin username and password, then provides an interface to unlock user accounts.

## Prerequisites

- .NET Framework installed
- System.DirectoryServices.AccountManagement
- Access to an Active Directory environment

## How to Use

1. Clone or download this repository.
2. Compile the program using Visual Studio or any C# compiler.
3. Run the compiled executable.
4. Enter the admin username and password when prompted.
5. Follow the instructions:
     > -   1- Unlock specific user.
     > -   2- Check if any accounts are locked.
     > -   3- Unlock all locked accounts.
     > -   4- Exit/Close Application
7. After unlocking accounts or exiting the program, press any key to close the application.

## Features

- **Unlock Specific User:** Allows unlocking a specific user account by entering the username.
- **Check Locked Accounts:** Displays all locked user accounts in Active Directory.
- **Unlock All Locked Accounts:** Unlocks all locked user accounts in Active Directory.

## Important Notes

- Ensure that you have the necessary permissions to unlock user accounts in Active Directory.
- This program does not store or log any credentials provided.
