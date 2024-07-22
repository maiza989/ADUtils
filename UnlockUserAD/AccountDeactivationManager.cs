using Pastel;
using System;
using System.DirectoryServices.AccountManagement;
using System.Drawing;

namespace ADUtils
{
    public class AccountDeactivationManager
    {
        public void DeactivateUserAccount(PrincipalContext context)
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
            } while(!returnToMenu);   
        }// end of DeactivateUserAccount
    }// end of class
}// end of namespace
