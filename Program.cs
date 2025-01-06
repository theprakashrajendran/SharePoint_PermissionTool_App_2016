using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using OfficeOpenXml.Style;
using System.Drawing;
using Microsoft.SharePoint.Client;
using System.Text;
using System.Net;
using System.Security;

namespace BASF.SharePoint.PermissionTool.Scheduler
{
    class SiteDetails
    {
        public string SiteUrl { get; set; }
        public string Period { get; set; }
        public string ReportUniuqeCode { get; set; }
    }
    class Program
    {
        #region VARIABLES
        public static string sAppName = System.Configuration.ConfigurationManager.AppSettings["APPNAME"].ToString();
        public static string dInternalUser = System.Configuration.ConfigurationManager.AppSettings["InternalUserDomain"].ToString();
        public static string dExternalUser = System.Configuration.ConfigurationManager.AppSettings["ExternalUserrDomain"].ToString();
        public static string dContractUser = System.Configuration.ConfigurationManager.AppSettings["ContractUserrDomain"].ToString();
        public static string dADUser = System.Configuration.ConfigurationManager.AppSettings["ADGroupUserDomain"].ToString();
        public static string sUserName = System.Configuration.ConfigurationManager.AppSettings["UserName"].ToString();
        public static string sPassword = System.Configuration.ConfigurationManager.AppSettings["Password"].ToString();
        public static string CheckSPO = System.Configuration.ConfigurationManager.AppSettings["CheckSPO"].ToString();
        public static List<SecurableObject> securableObjects;
        #endregion

        static void Main(string[] args)
        {
            Console.Title = sAppName;
            StartApplication();
        }

        #region APPLICATION START
        private static void StartApplication()
        {
            var lastdayOfThisWeek = DateTime.Now.LastDayOfWeek();
            var firstDayOfMonth = DateTime.Now.FirstDayOfMonth();
            SQL.InsertApplicationLog("Scheduler Started");
            DataTable _siteCollections = SQL.GetSiteCollections();
            /*DataTable _siteCollections = new DataTable();
            _siteCollections.Columns.Add("SID", typeof(int));
            _siteCollections.Columns.Add("Url");
            _siteCollections.Columns.Add("Period");
            _siteCollections.Columns.Add("ReportUniqueCode");
            _siteCollections.Rows.Add(1, "http://hisvshrptpoc002:5050/", "Monthly", "123");*/

            List<SiteDetails> oSiteDetails = new List<SiteDetails>();
            foreach (DataRow dr in _siteCollections.Rows)
            {
                oSiteDetails.Add(new SiteDetails
                {
                    SiteUrl = Convert.ToString(dr.ItemArray[1]),
                    Period = Convert.ToString(dr.ItemArray[2]),
                    ReportUniuqeCode = Convert.ToString(dr.ItemArray[3])
                });
            }
            DateTime scheduledDateWeekly = new DateTime(Convert.ToInt32(lastdayOfThisWeek.Year), Convert.ToInt32(lastdayOfThisWeek.Month), Convert.ToInt32(lastdayOfThisWeek.Day), 0, 0, 0);
            DateTime scheduledDateMonthly = new DateTime(Convert.ToInt32(firstDayOfMonth.Year), Convert.ToInt32(firstDayOfMonth.Month), Convert.ToInt32(firstDayOfMonth.Day), 0, 0, 0);

            DateTime todayDate = new DateTime(Convert.ToInt32(DateTime.Now.Year), Convert.ToInt32(DateTime.Now.Month), Convert.ToInt32(DateTime.Now.Day), 0, 0, 0);
            var oSchedule = new SiteDetails();
            oSiteDetails.RemoveAll(p => p.Period == "Weekly");
            oSiteDetails.RemoveAll(p => p.Period == "Daily");
            if (!CompareTwoDates(scheduledDateWeekly, todayDate))
            {

                //oSiteDetails.RemoveAll(p => p.Period == "Weekly");
            }
            if (!CompareTwoDates(scheduledDateMonthly, todayDate))
            {
                //oSiteDetails.RemoveAll(p => p.Period == "Monthly");
            }
            if (oSiteDetails.Count() > 0)
            {
                foreach (SiteDetails oSd in oSiteDetails)
                {
                    string url = Convert.ToString(oSd.SiteUrl);
                    string reportUniuqeCode = Convert.ToString(oSd.ReportUniuqeCode);
                    //SecurableObjectsReporting(url, reportUniuqeCode);
                    SecurableObjectsReporting(oSd);
                    //PrincipalsReporting(url, reportUniuqeCode);
                    //CrossSiteUsageReport(url, reportUniuqeCode);
                    //OrphanUserReport(url);
                }
            }

            Console.ForegroundColor = ConsoleColor.DarkYellow;
            SQL.InsertApplicationLog("Scheduler Completed");
            Console.WriteLine("Completed");
            Console.WriteLine();
            //Console.WriteLine("**************PRESS ANY KEY TO EXIT*****************");
            //Console.ReadLine();
        }
        #endregion

        #region Securable

        #region Securable Object Report

        public static bool SecurableObjectsReporting(SiteDetails siteDetails)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine();
                Console.WriteLine("*******************************");
                Console.WriteLine("Generating Securable Object Report");


                ReportTopInfo rptInfo = new ReportTopInfo();
                string strUsername = sUserName;
                string password = sPassword;
                SecureString secPwd = new SecureString();
                if (CheckSPO == DataMembers.CheckSPO)
                {
                    Array.ForEach(password.ToArray(), secPwd.AppendChar);
                    secPwd.MakeReadOnly();
                }

                using (ClientContext context = new ClientContext(siteDetails.SiteUrl))
                {
                    if (CheckSPO == DataMembers.CheckSPO)
                    {
                        context.Credentials = new SharePointOnlineCredentials(strUsername, secPwd);
                    }

                    Web web = context.Web;
                    Site site = context.Site;
                    List oList = context.Web.SiteUserInfoList;
                    CamlQuery q = new CamlQuery();
                    UserCollection userCollection = web.SiteUsers;
                    RoleAssignmentCollection rac = web.RoleAssignments;

                    securableObjects = new List<SecurableObject>();
                    //filter to get AD groups only
                    //q.ViewXml = "<View><Query><Where><Eq><FieldRef Name=\"ContentType\" /><Value Type=\"Text\">DomainGroup</Value></Eq></Where><OrderBy><FieldRef Name=\"Title\" /></OrderBy></Query></View>";
                    //Microsoft.SharePoint.Client.ListItemCollection collListItem = oList.GetItems(q);
                    context.Load(site, s => s.Owner, s => s.Usage);
                    context.Load(web);
                    context.Load(userCollection);
                    context.Load(web, website => website.Webs, website => website.Title, website => website.HasUniqueRoleAssignments, website => website.SiteGroups, website => website.Id);
                    //context.Load(rac, roleAssignement => roleAssignement.Include(r => r.Member, r => r.RoleDefinitionBindings));
                    //context.Load(collListItem,
                    //         items => items.Include(
                    //            item => item.Id,
                    //            item => item.DisplayName,
                    //            item => item.HasUniqueRoleAssignments,
                    //            item => item.ContentType));
                    context.ExecuteQuery();

                    string OwnersOfRootWeb = string.Empty;
                    foreach (User user in userCollection)
                    {
                        if (user.IsSiteAdmin)
                        {
                            OwnersOfRootWeb += user.Title + ';';
                        }
                    }
                    if (OwnersOfRootWeb.Contains(';'))
                    {
                        OwnersOfRootWeb = OwnersOfRootWeb.Remove(OwnersOfRootWeb.Length - 1, 1);
                    }

                    #region Headers

                    rptInfo.Title = "Security Report: Securable Objects";
                    rptInfo.CreatedDate = DateTime.Now.ToString("dd.MM.yyyy HH:mm");
                    rptInfo.SiteOwners = OwnersOfRootWeb;
                    rptInfo.SiteName = web.Title;
                    rptInfo.SiteUrl = web.Url;

                    #region Site Collection Usage Info

                    var siteUsage = site.Usage;
                    var storageQuota = siteUsage.Storage == 0 ? siteUsage.Storage.ToString() : (((siteUsage.Storage / siteUsage.StoragePercentageUsed) * 100) / (1024f * 1024f * 1024f)).ToString("f");
                    var used = siteUsage.Storage == 0 ? siteUsage.Storage : (siteUsage.Storage / (1024f * 1024f * 1024f));
                    var usagePercentage = siteUsage.StoragePercentageUsed;

                    #endregion

                    rptInfo.Usage = storageQuota + "GB/" + used.ToString("f") + "GB/" + usagePercentage.ToString("f") + "%";

                    #endregion

                    #region Get Securable Object Report

                    //SQL.InsertReportTopInformation(rptInfo, siteDetails.ReportUniuqeCode);

                    securableObjects = new List<SecurableObject>();
                    GetSecurableObjectReport(web, rptInfo, siteDetails, true);

                    #endregion

                    DataTable securableObjectDataTable = Common.ConvertToDataTable<SecurableObject>(securableObjects);

                    if (Common.ExportToExcel(securableObjectDataTable, REPORTNAMES.reportNames.SECURABLEOBJECTREPORT.ToString(), rptInfo, siteDetails.Period))
                    {
                        Console.WriteLine();
                        Console.WriteLine("Securable Object Report generated for site : " + rptInfo.SiteUrl);
                        Console.WriteLine();
                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine("Unable to generate Securable Object Report for site : " + rptInfo.SiteUrl);
                        Console.WriteLine();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Securable Object Report Generation Failed");
                SQL.InsertErrorLog("", sAppName, "Program.SecurableObjectsReporting", ex);
                return false;
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Generated Securable Object Report Successfully");
            return true;
        }

        #endregion

        private static void GetSecurableObjectReport(Web web, ReportTopInfo rptInfo, SiteDetails siteDetails, bool isRootWeb)
        {
            string Permission_Levels = string.Empty;
            try
            {

                if (CheckSiteTemplate(web))
                {
                    ClientContext context = new ClientContext(web.Url);
                    Web curWeb = context.Web;
                    context.Load(curWeb, thisWeb => thisWeb.Webs, thisWeb => thisWeb.Url, thisWeb => thisWeb.Title, thisWeb => thisWeb.HasUniqueRoleAssignments);
                    IEnumerable<Web> subWebs = context.LoadQuery(curWeb.Webs.Include(thisWeb => thisWeb.Title, thisWeb => thisWeb.WebTemplate, thisWeb => thisWeb.Webs, thisWeb => thisWeb.Url, thisWeb => thisWeb.HasUniqueRoleAssignments));

                    IEnumerable<RoleAssignment> roleAssignments = null;
                    IEnumerable<User> siteUsers = null;
                    IEnumerable<Group> groups = null;
                    if (isRootWeb || web.HasUniqueRoleAssignments)
                    {
                        roleAssignments = context.LoadQuery(
                            curWeb.RoleAssignments
                            .Include(
                                thisRole => thisRole.Member,
                                thisRole => thisRole.RoleDefinitionBindings));

                        siteUsers = context.LoadQuery(
                            context.Web.SiteUsers
                            .Include(
                                thisUser => thisUser.PrincipalType,
                                thisUser => thisUser.Email,
                                thisUser => thisUser.LoginName,
                                thisUser => thisUser.Title));

                        groups = context.LoadQuery(
                            context.Web.SiteGroups
                           .Include(
                               thisGroup => thisGroup.Users,
                               thisGroup => thisGroup.PrincipalType,
                               thisGroup => thisGroup.LoginName,
                               thisGroup => thisGroup.Title));
                    }

                    IEnumerable<List> lists = context.LoadQuery(
                        curWeb.Lists
                        .Where(thisList => !thisList.Hidden && !thisList.IsApplicationList && !thisList.IsPrivate)
                        .Include(
                            thisList => thisList.Title,
                            thisList => thisList.RootFolder,
                            thisList => thisList.HasUniqueRoleAssignments,
                            thisList => thisList.DefaultViewUrl,
                            thisList => thisList.RoleAssignments
                            .Include(
                                thisRole => thisRole.Member,
                                thisRole => thisRole.RoleDefinitionBindings)));
                    context.ExecuteQuery();

                    if (isRootWeb || web.HasUniqueRoleAssignments)
                    {
                        #region Group Details

                        if (groups != null && roleAssignments != null)
                        {
                            foreach (Group group in groups)
                            {
                                Permission_Levels = string.Empty;
                                foreach (var ra in roleAssignments.Where(curRole => curRole.Member.LoginName.Equals(group.LoginName)))
                                {
                                    foreach (var definition in ra.RoleDefinitionBindings)
                                    {
                                        if (!definition.Name.Equals("Limited Access"))
                                        {
                                            Permission_Levels += definition.Name + ';';
                                        }
                                    }
                                }

                                if (Permission_Levels.Equals(string.Empty))
                                {
                                    continue;
                                }

                                SecurableObject newObj = new SecurableObject();
                                newObj.ParentWebUrl = rptInfo.SiteUrl;
                                newObj.ParentName = rptInfo.SiteName;
                                newObj.ObjectUrl = curWeb.Url;
                                newObj.ObjectName = curWeb.Title;
                                newObj.ObjectType = ObjectType.Web;
                                newObj.PrincipalName = group.Title;
                                newObj.PrincipalEmail = string.Empty;
                                newObj.PrincipalAlias = group.LoginName;
                                newObj.PrincipalType = Convert.ToString(group.PrincipalType);

                                if (group.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.DistributionList) ||
                                    group.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.SecurityGroup))
                                {
                                    newObj.PrincipalSource = PrincipalSource.Groups.ActiveDirectory;
                                }
                                else if (group.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup))
                                {
                                    newObj.PrincipalSource = PrincipalSource.Groups.SharePointGroup;
                                }
                                newObj.PermissionLevel = Permission_Levels;
                                newObj.MemberCount = group.Users.Count;

                                securableObjects.Add(newObj);
                            }
                        }

                        #endregion

                        #region Site User Details

                        if (siteUsers != null && roleAssignments != null)
                        {
                            foreach (User user in siteUsers)
                            {
                                Permission_Levels = string.Empty;

                                foreach (var curRoleAssignment in roleAssignments.Where(ra => ra.Member.LoginName.Equals(user.LoginName)))
                                {
                                    foreach (var definition in curRoleAssignment.RoleDefinitionBindings)
                                    {
                                        if (!definition.Name.Equals("Limited Access"))
                                        {
                                            Permission_Levels += definition.Name + ';';
                                        }
                                    }
                                }

                                if (Permission_Levels.Equals(string.Empty))
                                {
                                    continue;
                                }

                                SecurableObject newObj = new SecurableObject();
                                newObj.ParentWebUrl = rptInfo.SiteUrl;
                                newObj.ParentName = rptInfo.SiteName;
                                newObj.ObjectUrl = curWeb.Url;
                                newObj.ObjectName = curWeb.Title;
                                newObj.ObjectType = ObjectType.Web;
                                newObj.PrincipalName = user.Title;
                                newObj.PrincipalEmail = user.Email;
                                newObj.PrincipalAlias = user.LoginName;
                                newObj.PrincipalType = PrincipalType.User;

                                if (user.LoginName.Equals("c:0(.s|true") ||
                                    user.LoginName.Equals("c:0!.s|windows") ||
                                    user.LoginName.Equals("SHAREPOINT\\system") ||
                                    user.LoginName.Equals("NT AUTHORITY\\LOCAL SERVICE"))
                                {
                                    newObj.PrincipalSource = PrincipalSource.Groups.BroadAccessRoles;
                                }
                                else if (user.Email.EndsWith("@basf.com"))
                                {
                                    newObj.PrincipalSource = PrincipalSource.User.Employee;
                                }
                                else if (user.Email.EndsWith("@partner.basf.com"))
                                {
                                    newObj.PrincipalSource = PrincipalSource.User.Partner;
                                }
                                else if (user.Email.EndsWith("@external-basf.com"))
                                {
                                    newObj.PrincipalSource = PrincipalSource.User.ExternalUser;
                                }
                                else
                                {
                                    newObj.PrincipalSource = PrincipalSource.User.UntrustedUser;
                                }
                                newObj.PermissionLevel = Permission_Levels;
                                newObj.MemberCount = 0;

                                securableObjects.Add(newObj);

                            }
                        }

                        #endregion
                    }

                    #region List Details

                    foreach (List curList in lists)
                    {
                        if (curList.HasUniqueRoleAssignments)
                        {
                            foreach (RoleAssignment curRoleAssignment in curList.RoleAssignments)
                            {
                                Permission_Levels = string.Empty;
                                foreach (var definition in curRoleAssignment.RoleDefinitionBindings)
                                {
                                    if (!definition.Name.Equals("Limited Access"))
                                    {
                                        Permission_Levels += definition.Name + ';';
                                    }
                                }

                                if (Permission_Levels.Equals(string.Empty))
                                {
                                    continue;
                                }

                                SecurableObject newObj = new SecurableObject();
                                newObj.ParentWebUrl = rptInfo.SiteUrl;
                                newObj.ParentName = rptInfo.SiteName;
                                newObj.ObjectUrl = curWeb.Url + curList.RootFolder.ServerRelativeUrl;
                                newObj.ObjectName = curList.Title;
                                newObj.ObjectType = ObjectType.List;
                                newObj.PrincipalName = curRoleAssignment.Member.Title;
                                if (curRoleAssignment.Member.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.User))
                                {
                                    newObj.PrincipalEmail = ((User)curRoleAssignment.Member).Email;
                                }
                                newObj.PrincipalAlias = curRoleAssignment.Member.LoginName;
                                newObj.PrincipalType = Convert.ToString(curRoleAssignment.Member.PrincipalType);

                                if (curRoleAssignment.Member.LoginName.Equals("c:0(.s|true") ||
                                    curRoleAssignment.Member.LoginName.Equals("c:0!.s|windows") ||
                                    curRoleAssignment.Member.LoginName.Equals("SHAREPOINT\\system") ||
                                    curRoleAssignment.Member.LoginName.Equals("NT AUTHORITY\\LOCAL SERVICE"))
                                {
                                    newObj.PrincipalSource = PrincipalSource.Groups.BroadAccessRoles;
                                }
                                else if (curRoleAssignment.Member.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.DistributionList) ||
                                    curRoleAssignment.Member.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.SecurityGroup))
                                {
                                    newObj.PrincipalSource = PrincipalSource.Groups.ActiveDirectory;
                                }
                                else if (curRoleAssignment.Member.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup))
                                {
                                    newObj.PrincipalSource = PrincipalSource.Groups.SharePointGroup;
                                    Group curGroup = (Group)curRoleAssignment.Member;
                                    context.Load(curGroup, thisGroup => thisGroup.Users);
                                    context.ExecuteQuery();
                                    newObj.MemberCount = curGroup.Users.Count;
                                }
                                else if (curRoleAssignment.Member.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.User))
                                {
                                    User user = ((User)curRoleAssignment.Member);
                                    if (user.Email.EndsWith("@basf.com"))
                                    {
                                        newObj.PrincipalSource = PrincipalSource.User.Employee;
                                    }
                                    else if (user.Email.EndsWith("@partner.basf.com"))
                                    {
                                        newObj.PrincipalSource = PrincipalSource.User.Partner;
                                    }
                                    else if (user.Email.EndsWith("@external-basf.com"))
                                    {
                                        newObj.PrincipalSource = PrincipalSource.User.ExternalUser;
                                    }
                                    else
                                    {
                                        newObj.PrincipalSource = PrincipalSource.User.UntrustedUser;
                                    }
                                }

                                newObj.PermissionLevel = Permission_Levels;

                                securableObjects.Add(newObj);
                            }
                        }

                        #region List Item Details

                        ListItemCollection listItemColl = curList.GetItems(CamlQuery.CreateAllItemsQuery());
                        IEnumerable<ListItem> listItems = context.LoadQuery(
                            listItemColl
                            .Where(thisListItem => thisListItem.HasUniqueRoleAssignments)
                            .Include(
                                thisListItem => thisListItem.Id,
                                thisListItem => thisListItem.HasUniqueRoleAssignments,
                                thisListItem => thisListItem.FieldValuesAsText,
                                thisListItem => thisListItem.FileSystemObjectType,
                                thisListItem => thisListItem.ParentList.DefaultDisplayFormUrl,
                                thisListItem => thisListItem.RoleAssignments
                                    .Include(
                                        curRole => curRole.Member,
                                        curRole => curRole.RoleDefinitionBindings)));
                        context.ExecuteQuery();

                        foreach (ListItem curListItem in listItems)
                        {
                            if (curListItem.HasUniqueRoleAssignments)
                            {
                                foreach (RoleAssignment curRoleAssignment in curListItem.RoleAssignments)
                                {
                                    Permission_Levels = string.Empty;
                                    foreach (var definition in curRoleAssignment.RoleDefinitionBindings)
                                    {
                                        if (!definition.Name.Equals("Limited Access"))
                                        {
                                            Permission_Levels += definition.Name + ';';
                                        }
                                    }

                                    if (Permission_Levels.Equals(string.Empty))
                                    {
                                        continue;
                                    }
                                    SecurableObject newObj = new SecurableObject();
                                    newObj.ParentWebUrl = rptInfo.SiteUrl;
                                    newObj.ParentName = rptInfo.SiteName;
                                    if (curListItem.FileSystemObjectType.Equals(FileSystemObjectType.Folder))
                                    {
                                        newObj.ObjectType = ObjectType.Folder;
                                        context.Load(curListItem.Folder);
                                        context.ExecuteQuery();
                                        newObj.ObjectUrl = curWeb.Url + curListItem.Folder.ServerRelativeUrl;
                                    }
                                    else
                                    {
                                        newObj.ObjectType = ObjectType.Item;
                                        newObj.ObjectUrl = curWeb.Url + curListItem.ParentList.DefaultDisplayFormUrl + "?ID=" + curListItem.Id;
                                    }
                                    newObj.ObjectName = curListItem.FieldValuesAsText["Title"];
                                    newObj.PrincipalName = curRoleAssignment.Member.Title;

                                    if (curRoleAssignment.Member.LoginName.Equals("c:0(.s|true") ||
                                        curRoleAssignment.Member.LoginName.Equals("c:0!.s|windows") ||
                                        curRoleAssignment.Member.LoginName.Equals("SHAREPOINT\\system") ||
                                        curRoleAssignment.Member.LoginName.Equals("NT AUTHORITY\\LOCAL SERVICE"))
                                    {
                                        newObj.PrincipalSource = PrincipalSource.Groups.BroadAccessRoles;
                                    }
                                    else if (curRoleAssignment.Member.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.User))
                                    {
                                        newObj.PrincipalEmail = ((User)curRoleAssignment.Member).Email;
                                    }
                                    newObj.PrincipalAlias = curRoleAssignment.Member.LoginName;
                                    newObj.PrincipalType = Convert.ToString(curRoleAssignment.Member.PrincipalType);
                                    if (curRoleAssignment.Member.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.DistributionList) ||
                                        curRoleAssignment.Member.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.SecurityGroup))
                                    {
                                        newObj.PrincipalSource = PrincipalSource.Groups.ActiveDirectory;
                                    }
                                    else if (curRoleAssignment.Member.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.SharePointGroup))
                                    {
                                        newObj.PrincipalSource = PrincipalSource.Groups.SharePointGroup;
                                        Group curGroup = (Group)curRoleAssignment.Member;
                                        context.Load(curGroup, thisGroup => thisGroup.Users);
                                        context.ExecuteQuery();
                                        newObj.MemberCount = curGroup.Users.Count;
                                    }
                                    else if (curRoleAssignment.Member.PrincipalType.Equals(Microsoft.SharePoint.Client.Utilities.PrincipalType.User))
                                    {
                                        User user = ((User)curRoleAssignment.Member);
                                        if (user.Email.EndsWith("@basf.com"))
                                        {
                                            newObj.PrincipalSource = PrincipalSource.User.Employee;
                                        }
                                        else if (user.Email.EndsWith("@partner.basf.com"))
                                        {
                                            newObj.PrincipalSource = PrincipalSource.User.Partner;
                                        }
                                        else if (user.Email.EndsWith("@external-basf.com"))
                                        {
                                            newObj.PrincipalSource = PrincipalSource.User.ExternalUser;
                                        }
                                        else
                                        {
                                            newObj.PrincipalSource = PrincipalSource.User.UntrustedUser;
                                        }
                                    }

                                    newObj.PermissionLevel = Permission_Levels;

                                    securableObjects.Add(newObj);
                                }
                            }
                        }

                        #endregion
                    }

                    #endregion

                    foreach (Web subWeb in subWebs)
                    {
                        GetSecurableObjectReport(subWeb, rptInfo, siteDetails, false);
                    }
                }
            }
            catch (Exception ex)
            {
                SQL.InsertErrorLog("", sAppName, "Program.SecurableObjectsReporting.Get Subsite Info", ex);
            }
        }

        #endregion

        #region Prinipal Report
        #region Principal
        public static bool PrincipalsReporting(string _url, string _reportUniuqeCode)
        {
            try
            {
                Console.WriteLine();
                Console.WriteLine("*******************************");
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Generating Principal Report");

                DataMembers _members = new DataMembers();
                ReportTopInfo reportTopdata = new ReportTopInfo();
                List<GroupInfo> grpInfo = new List<GroupInfo>();

                var dataTable = GeneratePrincipalsReportingDataTable();
                using (ClientContext context = new ClientContext(_url))
                {
                    Web web = context.Web;
                    Site site = context.Site;
                    List oList = context.Web.SiteUserInfoList;
                    CamlQuery q = new CamlQuery();
                    UserCollection userCollection = web.SiteUsers;
                    //filter to get AD groups only
                    q.ViewXml = "<View><Query><Where><Eq><FieldRef Name=\"ContentType\" /><Value Type=\"Text\">DomainGroup</Value></Eq></Where><OrderBy><FieldRef Name=\"Title\" /></OrderBy></Query></View>";
                    Microsoft.SharePoint.Client.ListItemCollection collListItem = oList.GetItems(q);
                    context.Load(site, s => s.Owner);
                    context.Load(web);
                    context.Load(userCollection);
                    context.Load(web, website => website.Webs, website => website.Title, website => website.HasUniqueRoleAssignments);
                    context.Load(collListItem,
                             items => items.Include(
                                item => item.Id,
                                item => item.DisplayName,
                                item => item.HasUniqueRoleAssignments,
                                item => item.ContentType));
                    context.ExecuteQuery();

                    string RootWeb = web.Title;
                    if (collListItem.Count > 0)
                    {
                        _members.Is_shared_with_large_audience = "Yes";
                    }
                    else
                    {
                        _members.Is_shared_with_large_audience = "No";
                    }
                    string OwnersOfRootWeb = string.Empty;

                    #region Site Collection Usage Info
                    context.Load(site, s => s.Usage);
                    context.ExecuteQuery();
                    var siteUsage = site.Usage;
                    var storageQuota = siteUsage.Storage == 0 ? siteUsage.Storage.ToString() : (((siteUsage.Storage / siteUsage.StoragePercentageUsed) * 100) / (1024f * 1024f * 1024f)).ToString("f");
                    var used = siteUsage.Storage == 0 ? siteUsage.Storage : (siteUsage.Storage / (1024f * 1024f * 1024f));
                    var usagePercentage = siteUsage.StoragePercentageUsed;
                    #endregion

                    #region Headers


                    foreach (User _user in userCollection)
                    {
                        if (_user.IsSiteAdmin)
                        {
                            OwnersOfRootWeb += _user.Title + ';';
                        }
                    }

                    if (OwnersOfRootWeb.Contains(';'))
                    {
                        OwnersOfRootWeb = OwnersOfRootWeb.Remove(OwnersOfRootWeb.Length - 1, 1);
                    }


                    reportTopdata.Title = "Security Report: Principals";
                    // reportTopdata.SiteUrl = web.Url;
                    //reportTopdata.SiteName = web.Title;
                    reportTopdata.SiteOwners = OwnersOfRootWeb;
                    reportTopdata.CreatedDate = DateTime.Now.ToString("dd.MM.yyyy HH:mm");
                    reportTopdata.Usage = storageQuota + "GB/" + used.ToString("f") + "GB/" + usagePercentage.ToString("f") + "%";

                    bool havCollectedCustomGroupInfo = true;
                    int i = 0;

                    GetPrincipal(context, web, web, _members, grpInfo, reportTopdata, userCollection, OwnersOfRootWeb, RootWeb, dataTable, _reportUniuqeCode, havCollectedCustomGroupInfo, i);
                    foreach (Web subWeb in web.Webs)
                    {
                        GetPrincipal(context, web, subWeb, _members, grpInfo, reportTopdata, userCollection, OwnersOfRootWeb, RootWeb, dataTable, _reportUniuqeCode, havCollectedCustomGroupInfo, i);
                    }
                    #endregion

                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Generating Principal Report Failed");
                SQL.InsertErrorLog("", sAppName, "Program.PrincipalsReporting", ex);
                return false;
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Principal Report Generated Successfully");
            return true;

        }
        #endregion

        private static DataTable GetPrincipal(ClientContext context, Web web, Web Subweb, DataMembers _members, List<GroupInfo> grpInfo, ReportTopInfo reportTopdata, UserCollection userCollection, string OwnersOfRootWeb, string RootWeb, DataTable dataTable, string _reportUniuqeCode, bool havCollectedCustomGroupInfo, int i)
        {

            #region Get All Groups in the Site Collection

            if (CheckSiteTemplate(Subweb))
            {
                reportTopdata.SiteUrl = Subweb.Url;
                reportTopdata.SiteName = Subweb.Title;
                context.Load(Subweb, s => s.AssociatedMemberGroup, s => s.AssociatedOwnerGroup, s => s.AssociatedVisitorGroup,
                s => s.HasUniqueRoleAssignments, s => s.RoleAssignments.Groups, s => s.SiteGroups);

                context.ExecuteQuery();

                grpInfo.Add(new GroupInfo()
                {
                    GroupName = Subweb.AssociatedVisitorGroup.Title,
                    AssociatedSiteId = Subweb.Id.ToString(),
                    AssociatedSiteUrl = Subweb.Url,
                    HasUniquePermission = Subweb.HasUniqueRoleAssignments,
                    SiteName = Subweb.Title
                });

                grpInfo.Add(new GroupInfo()
                {
                    GroupName = Subweb.AssociatedOwnerGroup.Title,
                    AssociatedSiteId = Subweb.Id.ToString(),
                    AssociatedSiteUrl = Subweb.Url,
                    HasUniquePermission = Subweb.HasUniqueRoleAssignments,
                    SiteName = Subweb.Title
                });

                grpInfo.Add(new GroupInfo()
                {
                    GroupName = Subweb.AssociatedMemberGroup.Title,
                    AssociatedSiteId = Subweb.Id.ToString(),
                    AssociatedSiteUrl = Subweb.Url,
                    HasUniquePermission = Subweb.HasUniqueRoleAssignments,
                    SiteName = Subweb.Title
                });

                if (havCollectedCustomGroupInfo)
                {
                    foreach (var customGrp in Subweb.RoleAssignments.Groups)
                    {
                        grpInfo.Add(new GroupInfo()
                        {
                            GroupName = customGrp.Title,
                            AssociatedSiteId = customGrp.Id.ToString(),
                            AssociatedSiteUrl = "",
                            HasUniquePermission = false,
                            SiteName = Subweb.Title
                        });
                    }


                }

                if (web.Webs.Count() - 1 == i)
                {
                    foreach (var customGrp in Subweb.SiteGroups)
                    {
                        var isExist = grpInfo.Where(t => t.GroupName == customGrp.Title).Any();
                        if (!isExist)
                        {
                            grpInfo.Add(new GroupInfo()
                            {
                                GroupName = customGrp.Title,
                                AssociatedSiteId = customGrp.Id.ToString(),
                                AssociatedSiteUrl = "",
                                HasUniquePermission = false,
                                SiteName = Subweb.Title
                            });
                        }
                    }
                }
                i++;

                havCollectedCustomGroupInfo = false;
                #region Get User Groups and Role Assignments
                foreach (User _user in userCollection)
                {
                    _members.Note = string.Empty;
                    _members.Row_Type = string.Empty;
                    _members.Object_Id = string.Empty;
                    _members.Object_Name = string.Empty;
                    _members.Object_Type = string.Empty;
                    _members.Parent_Id = string.Empty;
                    _members.Link_to_item = string.Empty;
                    _members.Has_unique_permissions = string.Empty;
                    _members.Is_shared_with_large_audience = string.Empty;
                    _members.Is_shared_with_contractors = string.Empty;
                    _members.Is_shared_with_externals = string.Empty;
                    _members.Principal_Name = string.Empty;
                    _members.Principal_Alias = string.Empty;
                    _members.Principal_Type = string.Empty;
                    _members.Permission_Level = string.Empty;
                    _members.Permission_Levels = string.Empty;
                    _members.Is_contractor = string.Empty;
                    _members.Is_external = string.Empty;
                    _members.Assigned_through_SP_Group = string.Empty;
                    _members.Is_nested_AD_Group = string.Empty;
                    _members.Parent_AD_Group = string.Empty;
                    _members.AD_Group_direct_members = string.Empty;



                    if (_user.IsSiteAdmin)
                    {
                        OwnersOfRootWeb += _user.Title + ';';
                    }
                    context.Load(_user.Groups);
                    context.ExecuteQuery();

                    _members.Row_Type = _user.PrincipalType.ToString();
                    _members.Principal_Alias = _user.LoginName.Contains("|") ? _user.LoginName.Split('|')[1].ToString() : _user.LoginName.ToString();
                    _members.Principal_Name = _user.Title;
                    _members.Principal_Type = "user";
                    _members.Is_external = "Need Information";
                    _members.Object_Id = Convert.ToString(web.Id);
                    _members.Object_Name = Convert.ToString(web.Title);
                    if (web.Title == RootWeb)
                    {
                        _members.Object_Type = "Root Site";
                    }
                    else
                    {
                        _members.Object_Type = "Sub Site";
                    }

                    if (web.HasUniqueRoleAssignments)
                    {
                        _members.Has_unique_permissions = "Yes";
                    }
                    else
                    {
                        _members.Has_unique_permissions = "No";
                    }

                    _members.Permission_Levels = string.Empty;

                    _members.Assigned_through_SP_Group = "Need to Identify";

                    _members.Is_nested_AD_Group = "false";

                    _members.Parent_AD_Group = "Need Information";

                    _members.Link_to_item = web.Url;

                    RoleAssignmentCollection roleAssignments = web.RoleAssignments;
                    foreach (var grp in _user.Groups)
                    {
                        #region For Testing
                        //if (_user.Title == "Vasanthkumar Murthy")
                        //{
                        //    foreach (var t in _user.Groups)
                        //    {
                        //        Console.WriteLine(t.Title);
                        //    }
                        //}
                        #endregion
                        context.Load(grp, groups => groups.Users, groups => groups.PrincipalType);
                        context.Load(roleAssignments, roleAssignement => roleAssignement.Include(r => r.Member, r => r.RoleDefinitionBindings));
                        context.ExecuteQuery();

                        foreach (var ra in roleAssignments)
                        {
                            context.Load(ra.Member);
                            context.Load(ra.RoleDefinitionBindings);
                            context.ExecuteQuery();

                            foreach (var definition in ra.RoleDefinitionBindings)
                            {
                                context.Load(definition);
                                context.ExecuteQuery();


                                var _grpObjectInfo = grpInfo.Select(t => t).Where(g => (g.GroupName == ra.Member.LoginName) || (g.GroupName == grp.Title)).FirstOrDefault();

                                _members.Object_Id = Convert.ToString(web.Id);
                                _members.Object_Name = Convert.ToString(web.Title);
                                if (web.Title == RootWeb)
                                {
                                    _members.Object_Type = "Root Site";
                                }
                                else
                                {
                                    _members.Object_Type = "Sub Site";
                                }


                                var userGrps = _user.Groups.Where(t => t.Title == grp.LoginName).Any();
                                _members.Permission_Levels = definition.Name;
                                if (ra.Member.PrincipalType.ToString().ToLower() != "sharepointgroup" && ra.Member.LoginName == _user.LoginName)
                                {
                                    if (_grpObjectInfo != null)
                                    {
                                        if (_grpObjectInfo.HasUniquePermission)
                                        {
                                            _members.Permission_Levels = GetUserPermissionofNonInheritedGroup(_grpObjectInfo.AssociatedSiteUrl, grp.Title);
                                            _members.Assigned_through_SP_Group = _grpObjectInfo.GroupName;
                                            _members.Object_Id = _grpObjectInfo.AssociatedSiteId;
                                            _members.Object_Name = _grpObjectInfo.SiteName;
                                            _members.Object_Type = "Sub Site";
                                        }
                                        else
                                        {
                                            _members.Assigned_through_SP_Group = "No";
                                        }
                                        dataTable.Rows.Add(_members.Note, reportTopdata.SiteName, reportTopdata.SiteUrl, _members.Row_Type, _members.Principal_Name,
                                                        _members.Principal_Alias, _members.Principal_Type,
                                                       _members.Is_external, _members.Object_Id, _members.Object_Name, _members.Object_Type, _members.Has_unique_permissions,
                                                       _members.Permission_Levels, _members.Assigned_through_SP_Group, _members.Is_nested_AD_Group,
                                                       _members.Parent_AD_Group, _members.Link_to_item, null);
                                    }

                                }
                                else if (ra.Member.LoginName == grp.Title)
                                {
                                    _members.Assigned_through_SP_Group = grp.Title;
                                    dataTable.Rows.Add(_members.Note, reportTopdata.SiteName, reportTopdata.SiteUrl, _members.Row_Type, _members.Principal_Name, _members.Principal_Alias, _members.Principal_Type,
                                                   _members.Is_external, _members.Object_Id, _members.Object_Name, _members.Object_Type, _members.Has_unique_permissions,
                                                   _members.Permission_Levels, _members.Assigned_through_SP_Group, _members.Is_nested_AD_Group,
                                                   _members.Parent_AD_Group, _members.Link_to_item, null);
                                }

                            }
                        }

                    }

                    if (_user.Groups.Count() == 0)
                    {
                        _members.Assigned_through_SP_Group = "No";
                        context.Load(roleAssignments, assignment => assignment.Include(role => role.Member, role => role.RoleDefinitionBindings));
                        context.ExecuteQuery();

                        foreach (var ra in roleAssignments)
                        {
                            var _grpObjectInfo = grpInfo.Select(t => t).Where(g => (g.GroupName == ra.Member.LoginName)).FirstOrDefault();
                            if (_grpObjectInfo != null)
                            {
                                if (_grpObjectInfo.HasUniquePermission)
                                {
                                    _members.Object_Id = _grpObjectInfo.AssociatedSiteId;
                                    _members.Object_Name = _grpObjectInfo.SiteName;
                                    _members.Object_Type = "Sub Site";
                                    _members.Assigned_through_SP_Group = _grpObjectInfo.GroupName;
                                }
                            }


                            dataTable.Rows.Add(_members.Note, reportTopdata.SiteName, reportTopdata.SiteUrl, _members.Row_Type, _members.Principal_Name,
                                       _members.Principal_Alias, _members.Principal_Type, _members.Is_external,
                                      _members.Object_Id, _members.Object_Name, _members.Object_Type, _members.Has_unique_permissions,
                                      _members.Permission_Levels, _members.Assigned_through_SP_Group, _members.Is_nested_AD_Group, _members.Parent_AD_Group,
                                      _members.Link_to_item, null);
                        }
                    }
                }

                #endregion

                DataSet set = new DataSet(web.Title);

                set.Tables.Add(dataTable);

                //    if (ExportToExcel(securableObjectDataTable, REPORTNAMES.reportNames.PRINCIPALREPORT.ToString(), reportTopdata))
                //   {
                DataTable dtToDataBase = GeneratePrincipalsReportingDataTable();
                foreach (DataRow dr in dataTable.Rows)
                {
                    //,[Created] == null (Set Default DateTime in DB)
                    dtToDataBase.Rows.Add("", dr["SiteName"].ToString(), dr["URL"].ToString(), dr["RowType"].ToString(), dr["PrincipalName"].ToString(),
                        dr["PrincipalAlias"], dr["PrincipalType"], dr["PrincipalCategory"], dr["ObjectId"], dr["ObjectName"],
                        dr["ObjectType"], dr["HasChildUniquePermissions"].ToString() == "Yes" ? true : false, dr["PermissionLevel"].ToString(), dr["AssignedThroughSPGroup"].ToString(),
                     dr["IsNestedADGroup"], dr["ParentADGroup"].ToString(), dr["ItemLink"].ToString(), DateTime.Now, _reportUniuqeCode);
                }

                if (InsertDataIntoDB(dtToDataBase, REPORTNAMES.reportNames.PRINCIPALREPORT.ToString()))
                {
                    string fileName = GetFileNameText(web.Title, REPORTTYPE.reportType.SecurableObject.ToString(), "Weekly");
                    bool result = SQL.SchedulerReport(web.Title, web.Url, REPORTTYPE.reportType.SecurableObject.ToString(), "Weekly", fileName, _reportUniuqeCode);
                    if (result)
                    {
                        SQL.SchedulerUpdate(_reportUniuqeCode, REPORTTYPE.reportType.PrincipalObject.ToString());
                    }
                }
                else
                {
                    throw new Exception();
                }
            }

            #endregion



            return dataTable;
        }

        #region  Generate Principal Report Data Table
        private static DataTable GeneratePrincipalsReportingDataTable()
        {
            DataTable table = new DataTable();
            //      [SID]
            //,[SiteName]
            //,[URL]
            //,[RowType]
            //,[PrincipalName]
            //,[PrincipalAlias]
            //,[PrincipalType]
            //,[PrincipalCategory]
            //,[ObjectId]
            //,[ObjectName]
            //,[ObjectType]
            //,[HasChildUniquePermissions]
            //,[PermissionLevel]
            //,[AssignedThroughSPGroup]
            //,[IsNestedADGroup]
            //,[ParentADGroup]
            //,[ItemLink]
            //,[Created]
            table.Columns.Add("SID");//Note
            table.Columns.Add("SiteName");
            table.Columns.Add("URL");
            table.Columns.Add("RowType");
            table.Columns.Add("PrincipalName");
            table.Columns.Add("PrincipalAlias");
            table.Columns.Add("PrincipalType");
            table.Columns.Add("PrincipalCategory");
            table.Columns.Add("ObjectId");
            table.Columns.Add("ObjectName");
            table.Columns.Add("ObjectType");
            table.Columns.Add("HasChildUniquePermissions");
            table.Columns.Add("PermissionLevel");
            table.Columns.Add("AssignedThroughSPGroup");
            table.Columns.Add("IsNestedADGroup");
            table.Columns.Add("ParentADGroup");
            table.Columns.Add("ItemLink");
            table.Columns.Add("Created");
            table.Columns.Add("ReportUniuqeCode");

            //table.Columns.Add("Note");//
            //table.Columns.Add("Row_Type");
            //table.Columns.Add("Principal_Name");
            //table.Columns.Add("Principal_Alias");
            //table.Columns.Add("Principal_Type");
            //table.Columns.Add("Is_internal,_contractor_or_external");
            //table.Columns.Add("Object_Id");
            //table.Columns.Add("Object_Name");
            //table.Columns.Add("Object_Type");
            //table.Columns.Add("Has_children_with_unique_permissions");
            //table.Columns.Add("Permission_Level");
            //table.Columns.Add("Assigned_through_SP_Group");//
            //table.Columns.Add("Is_nested_AD_Group");//
            //table.Columns.Add("Parent_AD_Group");//
            //table.Columns.Add("Link_to_item");

            return table;
        }
        #endregion

        #region Get Non Inherited Group Permission
        private static string GetUserPermissionofNonInheritedGroup(string siteUrl, string groupTitle)
        {
            var permissionLevel = string.Empty;
            try
            {
                using (ClientContext context = new ClientContext(siteUrl))
                {
                    Web web = context.Web;
                    RoleAssignmentCollection roleAssignments = web.RoleAssignments;
                    context.Load(roleAssignments, roleAssignement => roleAssignement.Include(r => r.Member, r => r.RoleDefinitionBindings));
                    context.ExecuteQuery();

                    foreach (var ra in roleAssignments)
                    {
                        context.Load(ra.Member);
                        context.Load(ra.RoleDefinitionBindings);
                        context.ExecuteQuery();

                        foreach (var definition in ra.RoleDefinitionBindings)
                        {
                            context.Load(definition);
                            context.ExecuteQuery();
                            if (ra.Member.LoginName == groupTitle)
                            {
                                permissionLevel = definition.Name;
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                SQL.InsertErrorLog("", sAppName, "Program.GetUserPermissionofNonInheritedGroup", ex);
            }
            return permissionLevel;
        }
        #endregion
        #endregion

        #region Cross Site Usage
        #region Cross Site Usage Report

        private static bool CrossSiteUsageReport(string _url, string _reportUniuqeCode)
        {
            try
            {
                Console.WriteLine();
                Console.WriteLine("*******************************");
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Generating Cross Site Usage Report");

                ReportTopInfo rptInfo = new ReportTopInfo();
                DataMembers _mem = new DataMembers();
                DataTable dataTable = GenerateCrossSiteReportingDataTable();
                //table = GetChildAtPoi
                rptInfo.Title = "Service Report: Cross Site Collection User Statistics and Usage";
                rptInfo.CreatedDate = DateTime.Now.ToString("dd.MM.yyyy HH:mm");
                using (ClientContext context = new ClientContext(_url))
                {
                    Web web = context.Web;
                    Site site = context.Site;
                    UserCollection userCollection = web.SiteUsers;

                    context.Load(site, s => s.Owner, s => s);
                    context.Load(web);
                    context.Load(userCollection);
                    context.Load(web, website => website.Webs, website => website.Title, website => website.HasUniqueRoleAssignments);


                    #region Site Collection Usage Info
                    context.Load(site, s => s.Usage);
                    context.ExecuteQuery();
                    var siteUsage = site.Usage;
                    var storageQuota = siteUsage.Storage == 0 ? siteUsage.Storage.ToString("f") + "GB" : (((siteUsage.Storage / siteUsage.StoragePercentageUsed) * 100) / (1024f * 1024f * 1024f)).ToString("f") + "GB";
                    var used = siteUsage.Storage == 0 ? siteUsage.Storage.ToString("f") + "GB" : (siteUsage.Storage / (1024f * 1024f * 1024f)).ToString("f") + "GB";
                    var usagePercentage = siteUsage.StoragePercentageUsed.ToString("f") + "%";
                    #endregion

                    rptInfo.SiteUrl = web.Url;
                    _mem.Note = string.Empty;
                    _mem.Object_Id = site.Id.ToString();
                    _mem.Link_to_item = site.Url;
                    _mem.Primary_Owner = site.Owner.Title;
                    //context.Load(site, s => s.SecondaryContact.UserId);
                    //context.ExecuteQuery();
                    #region GET SECONDARY CONTACT
                    context.Load(web, w => w.AllProperties);
                    context.ExecuteQuery();
                    //context.Load(site,s=>s.SecondaryContact);
                    //context.ExecuteQuery();
                    #endregion
                    _mem.Secondary_Owner = "Secondary Owner";
                    _mem.AdditionalUser = string.Empty;
                    _mem.Is_contractor = "0";
                    _mem.Is_external = "0";
                    _mem.Is_nested_AD_Group = "0";
                    _mem.internalUser = "0";
                    int exUser = 0, inUser = 0, adUser = 0, conUser = 0;
                    foreach (User _user in userCollection)
                    {
                        _mem.Row_Type = _user.PrincipalType.ToString();
                        _mem.Principal_Alias = _user.LoginName.Contains("|") ? _user.LoginName.Split('|')[1].ToString() : _user.LoginName.ToString();

                        if (_user.Email.Contains(dInternalUser))
                            inUser++;
                        if (_user.Email.Contains(dExternalUser))
                            exUser++;
                        if (_user.Email.Contains(dContractUser))
                            conUser++;
                        if (_user.Email.Contains(dADUser))
                            adUser++;
                    }
                    _mem.Is_contractor = conUser.ToString();
                    _mem.Is_external = exUser.ToString();
                    _mem.Is_nested_AD_Group = adUser.ToString();
                    _mem.internalUser = inUser.ToString();
                    _mem.date = DateTime.Now.ToString("MM/dd/yyyy HH:mm"); //Usage Information
                    rptInfo.SiteName = web.Title;

                    dataTable.Rows.Add(_mem.Note, _mem.Object_Id, rptInfo.SiteName, _mem.Link_to_item, _mem.Primary_Owner, _mem.Secondary_Owner, _mem.AdditionalUser,
                        storageQuota, used, usagePercentage, Convert.ToInt32(_mem.internalUser), Convert.ToInt32(_mem.Is_contractor), Convert.ToInt32(_mem.Is_external),
                        Convert.ToInt32(_mem.Is_nested_AD_Group), _mem.date, "");

                    //   if (ExportToExcel(securableObjectDataTable, REPORTNAMES.reportNames.STATISTICSANDUSAGEREPORT.ToString(), rptInfo))
                    //   {
                    DataTable dtToDataBase = GenerateCrossSiteReportingDataTable();
                    foreach (DataRow dr in dataTable.Rows)
                    {
                        //[Created] == null (Set Default DateTime in DB)
                        dtToDataBase.Rows.Add("", dr["Id"].ToString(), dr["SiteName"], dr["URL"].ToString(), dr["PrimaryOwnerAlias"].ToString(), dr["SecondaryOwnerAlias"].ToString(), dr["FullControlUsers"],
                         dr["Quota"], dr["StorageGB"].ToString(), dr["StoragePercentage"].ToString(), dr["InternalUsers"].ToString(), dr["Contractors"].ToString(),
                         dr["Externals"].ToString(), dr["NestedADGroupUsers"].ToString(), DateTime.ParseExact(Convert.ToString(dr["LastWriteAccess"]), "MM/dd/yyyy HH:mm", null), null);
                    }
                    if (InsertDataIntoDB(dtToDataBase, REPORTNAMES.reportNames.STATISTICSANDUSAGEREPORT.ToString()))
                    {
                        string fileName = GetFileNameText(web.Title, REPORTTYPE.reportType.SecurableObject.ToString(), "Weekly");
                        bool result = SQL.SchedulerReport(web.Title, web.Url, REPORTTYPE.reportType.SecurableObject.ToString(), "Weekly", fileName, _reportUniuqeCode);
                        if (result)
                        {
                            SQL.SchedulerUpdate(_reportUniuqeCode, REPORTTYPE.reportType.StorageQuota.ToString());
                        }
                    }
                    else
                    {
                        throw new Exception();
                    }

                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Generating Cross Site Usage Report Failed");
                SQL.InsertErrorLog("", sAppName, "Program.CrossSiteUsageReport", ex);
                return false;
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Cross Site Usage Report Generated Successfully");
            return true;
        }

        #endregion

        #region Generate Cross Site Report Data Table

        private static DataTable GenerateCrossSiteReportingDataTable()
        {
            DataTable table = new DataTable();

            table.Columns.Add("SID");
            table.Columns.Add("Id");
            table.Columns.Add("SiteName"); //Not in Excel
            table.Columns.Add("URL");
            table.Columns.Add("PrimaryOwnerAlias");
            table.Columns.Add("SecondaryOwnerAlias");
            table.Columns.Add("FullControlUsers");
            table.Columns.Add("Quota");
            table.Columns.Add("StorageGB");
            table.Columns.Add("StoragePercentage");
            table.Columns.Add("InternalUsers", typeof(int));
            table.Columns.Add("Contractors", typeof(int));
            table.Columns.Add("Externals", typeof(int));//
            table.Columns.Add("NestedADGroupUsers", typeof(int));
            table.Columns.Add("LastWriteAccess");
            table.Columns.Add("Created");//Not in Excel

            return table;
        }

        #endregion
        #endregion

        #region Orphan

        #region Generate Orphan User Report Data Table

        private static DataTable GenerateOrphanuserDataTable()
        {
            DataTable table = new DataTable();

            table.Columns.Add("SID");
            table.Columns.Add("RowType");
            table.Columns.Add("PrincipalName");
            table.Columns.Add("PrincipalAlias");
            table.Columns.Add("PrincipalCategory");
            table.Columns.Add("IsExpired");
            table.Columns.Add("IsDisabled");
            table.Columns.Add("Name");
            table.Columns.Add("URL");
            table.Columns.Add("PrimaryOwnerAlias");
            table.Columns.Add("SecondaryOwnerAlias");
            table.Columns.Add("Created");

            return table;
        }

        #endregion

        #region Orphan User Report

        private static bool OrphanUserReport(string _url)
        {
            try
            {
                Console.WriteLine();
                Console.WriteLine("*******************************");
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Generating Orphan User Report");

                ReportTopInfo rptInfo = new ReportTopInfo();
                DataMembers _mem = new DataMembers();
                DataTable dataTable = GenerateOrphanuserDataTable();
                //table = GetChildAtPoi
                rptInfo.Title = "Service Report: Inactive users with active permissions";
                rptInfo.CreatedDate = DateTime.Now.ToString("dd.MM.yyyy HH:mm");
                using (ClientContext context = new ClientContext(_url))
                {
                    Web web = context.Web;
                    Site site = context.Site;
                    UserCollection userCollection = web.SiteUsers;

                    context.Load(site, s => s.Owner, s => s);
                    context.Load(web);
                    context.Load(userCollection);
                    context.Load(web, website => website.Webs, website => website.Title, website => website.HasUniqueRoleAssignments);
                    context.ExecuteQuery();

                    rptInfo.SiteUrl = web.Url;
                    rptInfo.SiteName = web.Title;
                    _mem.Note = string.Empty;
                    _mem.Primary_Owner = site.Owner.Title;
                    #region GET SECONDARY CONTACT
                    context.Load(web, w => w.AllProperties);
                    context.ExecuteQuery();
                    #endregion
                    _mem.Secondary_Owner = "Secondary Owner";

                    foreach (User _user in userCollection)
                    {
                        _mem.Row_Type = _user.PrincipalType.ToString();
                        _mem.Principal_Alias = _user.LoginName.Contains("|") ? _user.LoginName.Split('|')[1].ToString() : _user.LoginName.ToString();
                        _mem.Principal_Name = _user.Title;
                        _mem.internalUser = "external"; //Need Information
                        string isDisabled = "FALSE"; //Need Information
                        string isExpired = "TRUE";//Need Information

                        dataTable.Rows.Add(_mem.Note, _mem.Row_Type, _mem.Principal_Name, _mem.Principal_Alias, _mem.internalUser, isExpired, isDisabled,
                            rptInfo.SiteName, rptInfo.SiteUrl, _mem.Primary_Owner, _mem.Secondary_Owner, DateTime.Now);
                    }



                    //  if (ExportToExcel(securableObjectDataTable, REPORTNAMES.reportNames.ORPHANUSERREPORT.ToString(), rptInfo))
                    //  {
                    DataTable dtToDataBase = GenerateOrphanuserDataTable();
                    foreach (DataRow dr in dataTable.Rows)
                    {
                        //,[Created] == null (Set Default DateTime in DB)
                        dtToDataBase.Rows.Add("", dr["RowType"].ToString(), dr["PrincipalName"].ToString(), dr["PrincipalAlias"].ToString(), dr["PrincipalCategory"].ToString(), dr["IsExpired"],
                         dr["IsDisabled"], dr["Name"].ToString(), dr["URL"].ToString(), dr["PrimaryOwnerAlias"].ToString(), dr["SecondaryOwnerAlias"].ToString(), null);
                    }

                    if (!InsertDataIntoDB(dtToDataBase, REPORTNAMES.reportNames.ORPHANUSERREPORT.ToString()))
                        throw new Exception();
                    //  }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Generating Orphan User Report Failed");
                SQL.InsertErrorLog("", sAppName, "Program.OrphanUserReport", ex);
                return false;
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Orphan User Report Generated Successfully");
            return true;
        }

        #endregion
        #endregion

        #region INSERT DATA TO DATABASE

        private static bool InsertDataIntoDB(DataTable dtTable, string reportType)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Uploading MetaData Information to DataBase");
                SQL.BulkWriteToDatabase(dtTable, reportType);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Failed to Upload MetaData Information to DataBase");
                SQL.InsertErrorLog("", sAppName, "Program.InsertDataIntoDB", ex);
                return false;
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Successfully Uploaded MetaData Information to DataBase");
            return true;
        }

        #endregion

        #region Is Check Large Audience
        private static string IsCheckLargeAudienceorAdGroup(ClientContext context, Web oWebsite)
        {
            var siteExist = false;
            var result = "No";
            RoleAssignmentCollection subroleAssignments = oWebsite.RoleAssignments;
            GroupCollection webGroups = oWebsite.SiteGroups;

            foreach (RoleAssignment grp in subroleAssignments)
            {
                if (grp.Member.PrincipalType.ToString() == "SecurityGroup" || grp.Member.PrincipalType.ToString() == "Domain Group")
                {
                    siteExist = true;
                }
            }
            if (!siteExist)
            {
                foreach (Group group in webGroups)
                {
                    if (group.Title.ToString() != "Excel Services Viewers")
                    {
                        context.Load(group, webGroup => webGroup.Title, webGroup => webGroup.Users);
                        context.ExecuteQuery();
                        foreach (User grpUser in group.Users)
                        {
                            if (grpUser.PrincipalType.ToString() == "SecurityGroup" || grpUser.PrincipalType.ToString() == "Domain Group")
                            {
                                result = "Yes";
                                break;
                            }
                        }
                    }
                }
            }
            return result;
        }
        #endregion

        #region check site template
        private static bool CheckSiteTemplate(Web web)
        {
            bool result = false;
            if (web.WebTemplate != "APP")
            {
                result = true;
            }
            return result;
        }
        #endregion

        public static string GenerateUniqueReportId(string reportType, string period)
        {
            string uniqueReportId = string.Empty;
            if (null != reportType)
            {
                string rt = GetReportType(reportType);
                StringBuilder builder = new StringBuilder();
                builder.Append("spprp");
                builder.Append(rt.ToLower());
                builder.Append(GetPeriodType(period).ToLower());
                builder.AppendFormat("{0:yyyyMdHHmmss}", DateTime.Now);
                uniqueReportId = builder.ToString();
            }
            return uniqueReportId.ToLower();
        }

        /// <summary>
        /// Method to generate filename for text file
        /// SiteName_ReportType_Period_Date
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string GetFileNameText(string siteName, string reportType, string period)
        {
            string fileName = siteName;
            if (null != siteName)
            {
                string rt = GetReportType(reportType);
                StringBuilder builder = new StringBuilder();
                builder.Append(siteName);
                builder.Append("_");
                builder.Append(rt);
                builder.Append("_");
                builder.Append(GetPeriodType(period));
                builder.AppendFormat("-{0:yyyy-M-d-HH-mm-ss}", DateTime.Now);
                fileName = builder.ToString();
            }
            return fileName;
        }

        private static string GetReportType(string reportType)
        {
            string result = "";
            if (reportType == REPORTTYPE.reportType.SecurableObject.ToString())
            {
                result = "SO";
            }
            else if (reportType == REPORTTYPE.reportType.PrincipalObject.ToString())
            {
                result = "PO";
            }
            else if (reportType == REPORTTYPE.reportType.StorageQuota.ToString())
            {
                result = "SQ";
            }
            return result;
        }

        private static string GetPeriodType(string periodType)
        {
            string result = "";
            if (periodType == PERIODTYPE.periodType.Daily.ToString())
            {
                result = "D";
            }
            else if (periodType == PERIODTYPE.periodType.Weekly.ToString())
            {
                result = "W";
            }
            else if (periodType == PERIODTYPE.periodType.Monthly.ToString())
            {
                result = "M";
            }
            return result;
        }
        private static bool CompareTwoDates(DateTime scheduledDate, DateTime actualDate)
        {
            bool relationship = false;
            int result = DateTime.Compare(scheduledDate, actualDate);
            if (result == 0)
            {
                relationship = true;
            }
            return relationship;
        }
    }
}
