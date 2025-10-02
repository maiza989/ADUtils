using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADUtils
{
    class TargetOU
    {
        public string Office { get; set; }
        public string Role { get; set; }
        public string ParentOU { get; set; }
        public string DisplayName { get; set; }

        public TargetOU(string office, string role, string parentOU, string displayName = null)
        {
            Office = office;
            Role = role;
            ParentOU = parentOU;
            DisplayName = displayName ?? role;
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
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "Accounting", "LM_Accounting", "NoAccountingEmail" }
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
        Groups = new List<string> { "_COLLECT", "CollectMI-11026982418", "_Training", "_Michigan", "Collectors", "LM_Collector", "NoOutboundEmail", "Horizon_Collector_RDS_Users", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "MI", Role = "Admin Staff",
        Groups = new List<string> { "_COLLECT", "CollectMI-11026982418", "_Training", "_Michigan", "Administrative", "Staff", "Horizon_RDS_Desktop_Users", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "MI", Role = "Atty",
        Groups = new List<string> { "_COLLECT", "CollectMI-11026982418", "_Training", "_Michigan", "Attorneys", "LM_Atty", "Horizon_Attorney_RDS_Users", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" }
    },
    new GroupAssignmentModel
    {
        Region = "MI", Role = "Acct",
        Groups = new List<string> { "_COLLECT", "CollectMI-11026982418", "_Training", "_Michigan", "Accounting", "ACHCC_Full", "Horizon_ACC_WycomeMI_Map", "Horizon_Accounting_RDS_Users", "MI_Accounting_Printers", "MI_All_Users_Printers", "BRP_Staff_Horizon_User" }
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
        Groups = new List<string> { "_COLLECT", "_Training", "CW_AllUsers", "_COLLECTGA", "LM_Accounting", "NoAccountingEmail", "Horizon_Accounting_RDS_Users", "BRP_Staff_Horizon_User" }
    },


    // KY Remote (optional if different from KY standard)
    new GroupAssignmentModel
    {
        Region = "KY-Remote", Role = "IT",
        Groups = new List<string> { "_COLLECT", "_COLLECTKY", "_Training", "IT", "LM_IT", "Horizon_IT_User" }
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
};
        public static List<string> GetGroups(string region, string role)
        {
            return groupAssignments
                .FirstOrDefault(g => g.Region.Equals(region, StringComparison.OrdinalIgnoreCase)
                                  && g.Role.Equals(role, StringComparison.OrdinalIgnoreCase))
                ?.Groups ?? new List<string>();
        }// end of GetGroups method 
    }// end of class
}// end of namespace
