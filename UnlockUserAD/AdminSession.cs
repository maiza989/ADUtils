using Microsoft.Extensions.Configuration;
using System.DirectoryServices.AccountManagement;
using System.Management.Automation;
using System.Security;

namespace ADUtils
{
    /// <summary>
    /// The credentials the operator authenticated with, held in one place for the session.
    ///
    /// Exists because privileged operations were binding to AD with
    /// <c>new PrincipalContext(ContextType.Domain)</c> -- no credentials -- which runs as whoever
    /// is logged into the workstation, not as the admin account that was typed in at startup.
    /// The audit log named the admin while AD recorded the interactive user. Every AD write should
    /// go through <see cref="CreateContext"/> so the two agree.
    ///
    /// Note: the *Exchange* path deliberately still uses the ambient identity. The interactive
    /// account holds modern Exchange RBAC (Organization Management); the admin account only has the
    /// legacy Exchange groups, so authenticating Exchange as the admin would break it.
    /// </summary>
    public static class AdminSession
    {
        private static string _password;

        /// <summary>sAMAccountName the operator authenticated with.</summary>
        public static string Username { get; private set; }

        /// <summary>Short domain name, e.g. "lmlawfirm" -- used to qualify the account for WinRM.</summary>
        public static string Domain { get; private set; }

        /// <summary>The server PrincipalContext binds to (config myDomainName -- a DC hostname).</summary>
        public static string DomainName { get; private set; }

        public static bool IsSet => !string.IsNullOrEmpty(Username) && !string.IsNullOrEmpty(_password);

        public static void Set(string username, string password, IConfiguration configuration)
        {
            Username = username;
            _password = password;
            Domain = configuration["AccountCreationSettings:myDomain"];
            DomainName = configuration["AccountCreationSettings:myDomainName"];
        }// end of Set

        /// <summary>
        /// A PrincipalContext bound as the admin account. Callers must dispose it.
        /// </summary>
        public static PrincipalContext CreateContext()
        {
            if (!IsSet)
            {
                // Better to fail loudly than to silently fall back to the interactive user, which is
                // the behaviour this class exists to remove.
                throw new InvalidOperationException("Admin credentials have not been set for this session.");
            }
            return new PrincipalContext(ContextType.Domain, DomainName, Username, _password);
        }// end of CreateContext

        /// <summary>
        /// The admin credentials as a PSCredential for PowerShell remoting.
        ///
        /// Qualifies the username with the domain, which WinRM requires, without double-qualifying
        /// a name the operator already typed as DOMAIN\user or a UPN.
        /// </summary>
        public static PSCredential CreatePsCredential()
        {
            if (!IsSet)
            {
                throw new InvalidOperationException("Admin credentials have not been set for this session.");
            }

            string user = Username.Contains('\\') || Username.Contains('@') || string.IsNullOrWhiteSpace(Domain)
                ? Username
                : $"{Domain}\\{Username}";

            SecureString secure = new SecureString();
            foreach (char c in _password) secure.AppendChar(c);
            secure.MakeReadOnly();

            return new PSCredential(user, secure);
        }// end of CreatePsCredential
    }// end of class
}// end of namespace
