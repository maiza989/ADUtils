using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADUtils
{
    class TargetOU
    {
        /// <summary>What the operator types at the office prompt (KY, MI, GA, REMOTE).</summary>
        public string Office { get; set; }
        public string Role { get; set; }
        public string ParentOU { get; set; }
        public string DisplayName { get; set; }

        /// <summary>
        /// The key used to look up group membership in <see cref="GroupAssignmentHelper"/>.
        /// Defaults to <see cref="Office"/>; supply it explicitly when the group table uses a
        /// different name than the office prompt (e.g. Office "REMOTE" -> Region "KY-Remote").
        /// </summary>
        public string Region { get; set; }

        /// <summary>
        /// The sub-OU under <see cref="ParentOU"/> that the account is moved into, or empty to
        /// place it directly in the parent.
        ///
        /// Separate from <see cref="Role"/> because the two genuinely differ: the AD OU is often
        /// plural or named differently from the role label ("Atty" -> OU=Attorneys,
        /// "GA_Staff" -> OU=Staff). Deriving the OU name from the role is what produced nine
        /// target paths that did not exist.
        /// </summary>
        public string SubOu { get; set; }

        public TargetOU(string office, string role, string parentOU, string displayName = null,
                        string region = null, string subOu = null)
        {
            Office = office;
            Role = role;
            ParentOU = parentOU;
            DisplayName = displayName ?? role;
            Region = region ?? office;

            // "General" has always meant "no sub-OU, sit in the parent".
            SubOu = subOu ?? (role == "General" ? string.Empty : role);
        }
    }// end of TragetOU class
    class GroupAssignmentModel
    {
        public string Region { get; set; }
        public string Role { get; set; }
        public List<string> Groups { get; set; }

    }// end of class
    public static class GroupAssignmentHelper
    {
        private static readonly List<GroupAssignmentModel> groupAssignments = new List<GroupAssignmentModel>
{
    // KY
    new GroupAssignmentModel
    {
        Region = "KY", Role = "IT",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "IT", "LM_IT" }
    },
    new GroupAssignmentModel
    {
        Region = "KY", Role = "Collector",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "Collectors", "LM_Collector", "NoOutboundEmail" }
    },
    new GroupAssignmentModel
    {
        Region = "KY", Role = "Admin Staff",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "Administrative", "Staff" }
    },
    new GroupAssignmentModel
    {
        Region = "KY", Role = "Atty",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "Attorneys", "LM_Atty" }
    },
    new GroupAssignmentModel
    {
        Region = "KY", Role = "Acct",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "Accounting", "LM_Accounting", "NoOutboundEmail-Accounting" }
    },
    new GroupAssignmentModel
    {
        Region = "KY", Role = "Compliance",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "Compliance" }
    },

    // MI
    new GroupAssignmentModel
    {
        Region = "MI", Role = "General",
        Groups = new List<string> { "_COLLECT", "_Training", "_Michigan", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "MI", Role = "Collector",
        Groups = new List<string> { "_COLLECT", "CollectMI", "_Training", "_Michigan", "Collectors", "LM_Collector", "NoOutboundEmail", "Horizon_Collector_RDS_Users", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "MI", Role = "Admin Staff",
        Groups = new List<string> { "_COLLECT", "CollectMI", "_Training", "_Michigan", "Administrative", "Staff", "Horizon_RDS_Desktop_Users", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "MI", Role = "Atty",
        Groups = new List<string> { "_COLLECT", "CollectMI", "_Training", "_Michigan", "Attorneys", "LM_Atty", "Horizon_Attorney_RDS_Users", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "MI", Role = "Acct",
        Groups = new List<string> { "_COLLECT", "CollectMI", "_Training", "_Michigan", "Accounting", "ACHCC_Full", "Horizon_ACC_WycomeMI_Map", "Horizon_Accounting_RDS_Users", "MI_Accounting_Printers", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" }
    },

    // GA
    new GroupAssignmentModel
    {
        Region = "GA", Role = "General",
        Groups = new List<string> { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA" }
    },
    new GroupAssignmentModel
    {
        Region = "GA", Role = "Collector",
        Groups = new List<string> { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA", "Collectors", "LM_Collector", "NoOutboundEmail", "Horizon_Collector_RDS_Users", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "GA", Role = "Admin Staff",
        Groups = new List<string> { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA", "Administrative", "Staff", "Horizon_RDS_Desktop_Users", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "GA", Role = "Atty",
        Groups = new List<string> { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA", "Attorneys", "LM_Atty", "Georgia attorneys", "Horizon_Attorney_RDS_Users", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "GA", Role = "Acct",
        Groups = new List<string> { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA", "LM_Accounting", "NoOutboundEmail-Accounting", "Horizon_Accounting_RDS_Users", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "GA", Role = "Call_Center",
        Groups = new List<string> { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA", "Collectors", "LM_Collector", "NoOutboundEmail", "Horizon_Collector_RDS_Users", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "GA", Role = "GA_Staff",
        Groups = new List<string> { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA", "Administrative", "Staff", "Horizon_RDS_Desktop_Users", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "GA", Role = "Accounting",
        Groups = new List<string> { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA", "LM_Accounting", "NoOutboundEmail-Accounting", "Horizon_Accounting_RDS_Users", "BRP_Staff_Horizon_User" }
    },

    // KY Remote (optional if different from KY standard)
    new GroupAssignmentModel
    {
        Region = "KY-Remote", Role = "IT",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "IT", "LM_IT", "Horizon_IT_Users" }
    },
    new GroupAssignmentModel
    {
        Region = "KY-Remote", Role = "Collector",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "Collectors", "LM_Collector", "NoOutboundEmail", "Horizon_Collector_RDS_Users" }
    },
    new GroupAssignmentModel
    {
        Region = "KY-Remote", Role = "Admin Staff",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "Administrative", "Staff", "Horizon_RDS_Desktop_Users" }
    },
    new GroupAssignmentModel
    {
        Region = "KY-Remote", Role = "Atty",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "Attorneys", "LM_Atty", "Horizon_Attorney_RDS_Users" }
    },
    new GroupAssignmentModel
    {
        Region = "KY-Remote", Role = "Acct",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "Accounting", "LM_Accounting", "NoOutboundEmail-Accounting", "Horizon_Accounting_RDS_Users" }
    },
    new GroupAssignmentModel
    {
        Region = "KY-Remote", Role = "Compliance",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "Compliance", "Horizon_RDS_Desktop_Users" }
    },
};
        public static List<string> GetGroups(string region, string role)
        {
            return groupAssignments
                .FirstOrDefault(g => g.Region.Equals(region, StringComparison.OrdinalIgnoreCase)
                                  && g.Role.Equals(role, StringComparison.OrdinalIgnoreCase))
                ?.Groups ?? new List<string>();
        }// end of GetGroups method

        /// <summary>
        /// Every (Region, Role) combination the group table defines, whether or not any
        /// <see cref="GetTargetOUs"/> row can actually select it. Used by the validation report to
        /// spot dead rows -- configuration that looks live but is unreachable.
        /// </summary>
        internal static IEnumerable<(string Region, string Role, int GroupCount)> GetAllAssignments()
        {
            return groupAssignments.Select(g => (g.Region, g.Role, g.Groups?.Count ?? 0));
        }// end of GetAllAssignments

        /// <summary>
        /// The office/role options offered during account creation.
        ///
        /// Lives here rather than as a local inside CreateUserAccount so that the validation report
        /// can check the same rows the creator uses. When it was a local, nothing could verify it,
        /// and it drifted until nine of its twenty-two target OUs no longer existed.
        /// </summary>
        internal static List<TargetOU> GetTargetOUs()
        {
            return new List<TargetOU>
            {
                // KY -- sub-OUs under LloydMc_Lou, verified to exist.
                new TargetOU("KY", "IT", "LloydMc_Lou"),
                new TargetOU("KY", "Collector", "LloydMc_Lou"),
                new TargetOU("KY", "Admin Staff", "LloydMc_Lou"),
                new TargetOU("KY", "Atty", "LloydMc_Lou"),
                new TargetOU("KY", "Acct", "LloydMc_Lou"),
                new TargetOU("KY", "Compliance", "LloydMc_Lou"),

                // MI -- only Attorneys has a sub-OU; everyone else sits in Michigan_Users itself.
                new TargetOU("MI", "General", "Michigan_Users", "Michigan Users"),
                new TargetOU("MI", "Collector", "Michigan_Users", "Michigan Collector", subOu: ""),
                new TargetOU("MI", "Admin Staff", "Michigan_Users", "Michigan Admin Staff", subOu: ""),
                new TargetOU("MI", "Atty", "Michigan_Users", "Michigan Atty", subOu: "Attorneys"),
                new TargetOU("MI", "Acct", "Michigan_Users", "Michigan Accounting", subOu: ""),

                // GA -- Cooling_Users was retired and replaced by Georgia_Users. The sub-OU names
                // there differ from the role labels, hence the explicit subOu values.
                new TargetOU("GA", "General", "Georgia_Users", "Default Georgia Users"),
                new TargetOU("GA", "Call_Center", "Georgia_Users", "Georgia Collector", subOu: "Call_Center"),
                new TargetOU("GA", "GA_Staff", "Georgia_Users", "Georgia Admin Staff", subOu: "Staff"),
                new TargetOU("GA", "Atty", "Georgia_Users", "Georgia Atty", subOu: "Attorneys"),
                new TargetOU("GA", "Accounting", "Georgia_Users", "Georgia Accounting", subOu: "Accounting"),

                // Remote. Office is what the operator types ("REMOTE"); Region must match the
                // Region values in the group table above ("KY-Remote") or the hire gets no groups.
                new TargetOU("REMOTE", "IT", "LloydMc_Lou", region: "KY-Remote"),
                new TargetOU("REMOTE", "Collector", "LloydMc_Lou", region: "KY-Remote"),
                new TargetOU("REMOTE", "Admin Staff", "LloydMc_Lou", region: "KY-Remote"),
                new TargetOU("REMOTE", "Atty", "LloydMc_Lou", region: "KY-Remote"),
                new TargetOU("REMOTE", "Acct", "LloydMc_Lou", region: "KY-Remote"),
                new TargetOU("REMOTE", "Compliance", "LloydMc_Lou", region: "KY-Remote")
            };
        }// end of GetTargetOUs

        /// <summary>
        /// The OU fragment an account is moved into, e.g. "OU=Attorneys,OU=Georgia_Users" or
        /// "OU=Michigan_Users" when there is no sub-OU.
        ///
        /// Shared by account creation and the validation report on purpose: if each built the path
        /// itself, the validator could pass while the creator targeted somewhere else entirely.
        /// </summary>
        internal static string BuildRelativeOuPath(TargetOU target)
        {
            return string.IsNullOrEmpty(target.SubOu)
                ? $"OU={target.ParentOU}"
                : $"OU={target.SubOu},OU={target.ParentOU}";
        }// end of BuildRelativeOuPath
    }// end of class
}// end of namespace
