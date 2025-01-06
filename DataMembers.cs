using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BASF.SharePoint.PermissionTool.Scheduler
{

    public class DataMembers
    {
        public const string RootWeb = "RootWeb";
        public const string SubWeb = "SubWeb";
        public const string ReportSchedulerCompleted = "Completed";
        public const string CheckSPO = "SPO";
        
        public string Object_Name { get; set; }

        public string Link_to_item { get; set; }

        public string Note { get; set; }

        public string Row_Type { get; set; }

        public string Object_Id { get; set; }

        public string Object_Type { get; set; }

        public string Parent_Id { get; set; }

        public string Has_unique_permissions { get; set; }

        public string Is_shared_with_large_audience { get; set; }

        public string Is_shared_with_contractors { get; set; }

        public string Is_shared_with_externals { get; set; }

        public string Principal_Name { get; set; }

        public string Principal_Alias { get; set; }

        public string Principal_Type { get; set; }

        public string Permission_Level { get; set; }

        public string Permission_Levels { get; set; }

        public string Is_contractor { get; set; }

        public string Is_external { get; set; }

        public string Assigned_through_SP_Group { get; set; }

        public string Is_nested_AD_Group { get; set; }

        public string Parent_AD_Group { get; set; }

        public string AD_Group_direct_members { get; set; }

        public string Secondary_Owner { get; set; }

        public string Primary_Owner { get; set; }

        public string AdditionalUser { get; set; }

        public string internalUser { get; set; }

        public string date { get; set; }
    }

    public class GroupInfo
    {
        public string GroupName { get; set; }

        public string AssociatedSiteUrl { get; set; }

        public string AssociatedSiteId { get; set; }

        public bool HasUniquePermission { get; set; }

        public string SiteName { get; set; }
    }

    public class ReportTopInfo
    {
        public string Title { get; set; }

        public string SiteUrl { get; set; }

        public string SiteName { get; set; }

        public string SiteOwners { get; set; }

        public string CreatedDate { get; set; }

        public string Usage { get; set; }
    }

    public static class REPORTNAMES
    {
        public enum reportNames
        {
            PRINCIPALREPORT,
            STATISTICSANDUSAGEREPORT,
            ORPHANUSERREPORT,
            SECURABLESTORAGEREPORT,
            SECURABLEOBJECTREPORT
        }
    }

    public static class REPORTTYPE
    {
        public enum reportType
        {
            PrincipalObject,
            SecurableObject,
            StorageQuota
        }
    }

    public static class PERIODTYPE
    {
        public enum periodType
        {
            Daily,
            Weekly,
            Monthly
        }
    }

    public class SecurableObject
    {
        public string ParentWebUrl { get; set; }

        public string ParentName { get; set; }

        public string ObjectUrl { get; set; }

        public string ObjectName { get; set; }

        public string ObjectType { get; set; }

        public string PrincipalName { get; set; }

        public string PrincipalEmail { get; set; }

        public string PrincipalAlias { get; set; }

        public string PrincipalType { get; set; }

        public string PrincipalSource { get; set; }

        public string PermissionLevel { get; set; }

        public int MemberCount { get; set; }
    }

    public static class PrincipalType
    {
        public const string User = "User";
        public const string SecurityGroup = "Security Group";
        public const string SharePointGroup = "SharePoint Group";
    }

    public static class PrincipalSource
    {
        public static class User
        {
            public const string Employee = "Employee";
            public const string Partner = "Partner";
            public const string ExternalUser = "External User";
            public const string UntrustedUser = "Untrusted User";
        }

        public static class Groups
        {
            public const string ActiveDirectory = "Active Directory";
            public const string BroadAccessRoles = "Broad Access Roles";
            public const string SharePointGroup = "SharePoint Group";
        }
    }

    public static class ObjectType
    {
        public const string Web = "Web";
        public const string List = "List";
        public const string Folder = "Folder";
        public const string Item = "Item";
    }
}