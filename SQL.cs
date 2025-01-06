using Microsoft.ApplicationBlocks.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace BASF.SharePoint.PermissionTool.Scheduler
{
    public static class SQL
    {
        #region VARIABLES
        public static string sConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["connectionString"].ToString();
        public static string sTableName_OrphanedUsers = System.Configuration.ConfigurationManager.AppSettings["TableName_OrphanedUsers"].ToString();
        public static string sTableName_Principals = System.Configuration.ConfigurationManager.AppSettings["TableName_Principals"].ToString();
        public static string sTableName_SecurableObjectsName = System.Configuration.ConfigurationManager.AppSettings["TableName_SecurableObjects"].ToString();
        public static string sTableName_Usage = System.Configuration.ConfigurationManager.AppSettings["TableName_Usage"].ToString();
        public static string sTableName_Errorlog = System.Configuration.ConfigurationManager.AppSettings["TableName_Errorlog"].ToString();
        public static string sTableName_Applicationlog = System.Configuration.ConfigurationManager.AppSettings["TableName_Applicationlog"].ToString();
        public static string sTableName_SiteCollections = System.Configuration.ConfigurationManager.AppSettings["TableName_SiteCollections"].ToString();

        #endregion

        #region SQL CONNECTION
        /// <summary>
        /// SQL CONNECTION STRING
        /// </summary>
        public static string ConnectionString = Convert.ToString(sConnectionString);

        /// <summary>
        /// SQL OPEN CONNECTION
        /// </summary>
        /// <returns></returns>
        private static SqlConnection OpenConnection()
        {
            SqlConnection myConn = null;
            try
            {
                myConn = new SqlConnection(ConnectionString);
                if (myConn.State == System.Data.ConnectionState.Closed || myConn.State != System.Data.ConnectionState.Connecting)
                {
                    myConn.Open();
                }
            }
            catch (Exception ex)
            {
                //Write to text file as SQL connection failed               
                return myConn;

            }
            return myConn;
        }
        #endregion

        #region BULK INSERT
        /// <summary>
        /// SQL BULK INSERT FOR REPORT DATA
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="reportType"></param>
        public static void BulkWriteToDatabase(DataTable dataTable, string reportType)
        {
            using (SqlConnection connection = OpenConnection())
            {
                SqlBulkCopy bulkCopy = new SqlBulkCopy(
                    connection,
                    SqlBulkCopyOptions.TableLock |
                    SqlBulkCopyOptions.FireTriggers |
                    SqlBulkCopyOptions.UseInternalTransaction,
                    null);

                // set the destination table name
                // bulkCopy.DestinationTableName = "BulkTest";
                bulkCopy.DestinationTableName = GetReportTableName(reportType);
                // write the data in the “dataTable”
                try
                {
                    //The Datatable column names should match the table column names
                    bulkCopy.WriteToServer(dataTable);
                }
                catch (Exception ex)
                {
                    connection.Close();
                    throw;
                }
                connection.Close();
            }

            dataTable.Clear();
        }
        #endregion

        #region GET REPORT TABLE NAME
        /// <summary>
        /// GET DATABASE REPORT TABLE NAME BASED ON THE PARAMETER
        /// </summary>
        /// <param name="reportType"></param>
        /// <returns></returns>
        private static string GetReportTableName(string reportType)
        {
            string tableName = string.Empty;
            try
            {
                if (!string.IsNullOrEmpty(reportType))
                {
                    if (reportType.ToUpper() == REPORTNAMES.reportNames.SECURABLEOBJECTREPORT.ToString())
                    {
                        tableName = Convert.ToString(sTableName_SecurableObjectsName);
                    }
                    else if (reportType.ToUpper() == REPORTNAMES.reportNames.ORPHANUSERREPORT.ToString())
                    {
                        tableName = Convert.ToString(sTableName_OrphanedUsers);
                    }
                    else if (reportType.ToUpper() == REPORTNAMES.reportNames.PRINCIPALREPORT.ToString())
                    {
                        tableName = Convert.ToString(sTableName_Principals);
                    }
                    else if (reportType.ToUpper() == REPORTNAMES.reportNames.STATISTICSANDUSAGEREPORT.ToString())
                    {
                        tableName = Convert.ToString(sTableName_Usage);
                    }
                }
                else
                {
                    //LOG report type us empty
                }
                return tableName;
            }
            catch (Exception ex)
            {
                return tableName;
                throw;
            }
        }
        #endregion

        #region Error Log
        public static int InsertErrorLog(string username, string application, string logsource, Exception ex)
        {
            Console.WriteLine(ex.Message);
            int rowsAffected;
            try
            {
                rowsAffected = SqlHelper.ExecuteNonQuery(sConnectionString, sTableName_Errorlog,
                    username, application, logsource, ErrorHandling.BuildMessage(ex));
            }
            catch (Exception ee)
            {
                throw;
            }
            return rowsAffected;
        }
        #endregion

        #region Application Log
        public static int InsertApplicationLog(string logMsg)
        {
            Console.WriteLine(logMsg);
            int rowsAffected;
            try
            {
                string query = "INSERT INTO " + sTableName_Applicationlog + " (RunAt, Message) VALUES (@RunAt, @Message)";
                List<SqlParameter> parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@RunAt", DateTime.Now.ToString()));
                parameters.Add(new SqlParameter("@Message", logMsg));
                rowsAffected = SqlHelper.ExecuteNonQuery(sConnectionString, CommandType.Text, query, parameters.ToArray());
            }
            catch (Exception ee)
            {
                throw;
            }
            return rowsAffected;
        }
        #endregion

        #region GET SITE COLLECTIONS
        public static DataTable GetSiteCollections()
        {
            DataTable dtable = new DataTable();
            try
            {
                string query = GetQuery("GetSiteCollections");
                DataSet ds = SqlHelper.ExecuteDataset(sConnectionString, CommandType.Text, query);
                if (ds != null)
                    dtable = ds.Tables[0];
            }
            catch (Exception ee)
            {
                throw;
            }
            return dtable;
        }
        #endregion

        /// <summary>
        /// Put entry into scheduler table
        /// </summary>
        /// <param name="SiteName"></param>
        /// <param name="SiteURL"></param>
        /// <param name="ReportType"></param>
        /// <param name="Period"></param>
        /// <param name="FileName"></param>
        /// <param name="ReportUniuqeCode"></param>
        /// <returns></returns>
        #region Scheduler
        public static bool SchedulerReport(string SiteName, string SiteURL, string ReportType, string Period, string FileName, string ReportUniuqeCode)
        {
            bool result = false;
            try
            {
                // define INSERT query with parameters
                string query = GetQuery("SchedulerReport");
                // create connection and command
                using (SqlConnection conn = new SqlConnection(sConnectionString))
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    // define parameters and their values
                    cmd.Parameters.Add("@SiteName", SqlDbType.NVarChar, 1000).Value = SiteName;
                    cmd.Parameters.Add("@SiteURL", SqlDbType.NVarChar, 1000).Value = SiteURL;
                    cmd.Parameters.Add("@ReportType", SqlDbType.NVarChar, 1000).Value = ReportType;
                    cmd.Parameters.Add("@Period", SqlDbType.NVarChar, 1000).Value = Period;
                    cmd.Parameters.Add("@IsAudit", SqlDbType.Bit).Value = true;
                    cmd.Parameters.Add("@IsDownloaded", SqlDbType.Bit).Value = true;
                    cmd.Parameters.Add("@IsActive", SqlDbType.Bit).Value = true;
                    cmd.Parameters.Add("@FileName", SqlDbType.NVarChar, 1000).Value = FileName;
                    cmd.Parameters.Add("@ReportUniuqeCode", SqlDbType.NVarChar, 1000).Value = ReportUniuqeCode;
                    cmd.Parameters.Add("@Created", SqlDbType.DateTime).Value = DateTime.Now;
                    cmd.Parameters.Add("@Modified", SqlDbType.DateTime).Value = DateTime.Now;
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    result = true;
                }
            }
            catch (Exception ex)
            {
                ErrorHandling.WriteLog("SchedulerReport", ex);
            }
            return result;
        }
        #endregion

        #region Scheduler table update
        public static bool SchedulerUpdate(string ReportUniuqeCode, string ReportType)
        {
            bool result = false;
            try
            {
                // define INSERT query with parameters
                string query = GetQuery(ReportType);
                // create connection and command
                using (SqlConnection conn = new SqlConnection(sConnectionString))
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    // define parameters and their values
                    cmd.Parameters.Add("@Modified", SqlDbType.DateTime).Value = DateTime.Now;
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    result = true;
                }
            }
            catch (Exception ex)
            {
                ErrorHandling.WriteLog("SchedulerUpdate", ex);
            }
            return result;
        }
        #endregion

        #region Insert Report Top Information
        public static bool InsertReportTopInformation(ReportTopInfo rptInfo, string ReportUniuqeCode)
        {
            bool result = false;
            try
            {
                // define INSERT query with parameters
                string query = GetQuery("InsertReportTopInformation");
                // create connection and command
                using (SqlConnection conn = new SqlConnection(sConnectionString))
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    // define parameters and their values
                    cmd.Parameters.Add("@Title", SqlDbType.NVarChar).Value = rptInfo.Title;
                    cmd.Parameters.Add("@SiteUL", SqlDbType.NVarChar).Value = rptInfo.SiteUrl;
                    cmd.Parameters.Add("@SiteOwners", SqlDbType.NVarChar).Value = rptInfo.SiteOwners;
                    cmd.Parameters.Add("@SiteCollUsage", SqlDbType.NVarChar).Value = rptInfo.Usage;
                    // cmd.Parameters.Add("@StorageQuota", SqlDbType.NVarChar).Value = rptInfo.;
                    cmd.Parameters.Add("@ReportUniuqeCode", SqlDbType.NVarChar).Value = ReportUniuqeCode;
                    cmd.Parameters.Add("@Created", SqlDbType.DateTime).Value = DateTime.Now;
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    result = true;
                }
            }
            catch (Exception ex)
            {
                ErrorHandling.WriteLog("SchedulerUpdate", ex);
            }
            return result;
        }
        #endregion

        #region GetUpdateQuery
        private static string GetQuery(string type)
        {
            string query = "";
            if (type == REPORTTYPE.reportType.SecurableObject.ToString())
            {
                query = "update SiteCollections set IsSecurable = 0,Modified=@Modified where ReportUniuqeCode = @ReportUniuqeCode and IsActive = 1";
            }
            else if (type == REPORTTYPE.reportType.PrincipalObject.ToString())
            {
                query = "update SiteCollections set IsPrincipal = 0,Modified=@Modified where ReportUniuqeCode = @ReportUniuqeCode and IsActive = 1";
            }
            else if (type == REPORTTYPE.reportType.StorageQuota.ToString())
            {
                query = "update SiteCollections set IsStorageQuota = 0,Modified=@Modified where ReportUniuqeCode = @ReportUniuqeCode and IsActive = 1";
            }
            else if (type == DataMembers.ReportSchedulerCompleted)
            {
                query = "update SiteCollections set IsActive = 0,Modified=@Modified where ReportUniuqeCode = @ReportUniuqeCode and IsActive = 1";
            }
            else if (type == "SchedulerReport")
            {
                query = "insert into Scheduler(SiteName, SiteURL, ReportType, Period, IsAudit, IsDownloaded, IsActive, FileName, ReportUniuqeCode, Created, Modified) values (@SiteName, @SiteURL, @ReportType, @Period, @IsAudit,@IsDownloaded, @IsActive, @FileName, @ReportUniuqeCode, @created, @Modified)";
            }
            else if (type == "GetSiteCollections")
            {
                query = "select SiteName,SiteURL,Period,ReportUniuqeCode,IsActive FROM " + sTableName_SiteCollections + " where IsActive = 1";
            }
            else if (type == "InsertReportTopInformation")
            {
                query = @"INSERT INTO  [ReportTopInfo] ([Title],[SiteURL] ,[SiteOwners] ,[SiteCollUsage],[ReportUniuqeCode] ,[Created]) VALUES   
                        (@Title , @SiteUL, @SiteOwners , @SiteCollUsage , @ReportUniuqeCode  , @Created)";
            }
            return query;
        }
        #endregion

        #region SECURABLE OBJECT FILE

        /// <summary>
        /// Function to store the data as BLOB (Excel file) in SQL
        /// </summary>
        /// <param name="byteData">File data in bytes</param>
        /// <param name="fileName">Name of the file along with extension</param>
        /// <param name="reportTopInfo">Report information</param>
        /// <param name="reportType">Report Type</param>
        /// <param name="period">Period on which the report is generated</param>
        /// <param name="isAudit">Is the file is for Audit?</param>
        /// <param name="isActive">Is the file is Active?</param>
        /// <param name="reportUniqueCode">Report unique code</param>
        /// <returns>1 if the file is stored and -1 if the file is not stored</returns>
        public static int InsertSecurableObjectFile(byte[] byteData, ReportTopInfo reportTopInfo, string reportType, string period, bool isAudit, bool isActive, string reportUniqueCode)
        {
            int rowsAffected = -1;
            try
            {
                string strQuery = @"INSERT INTO [dbo].[Scheduler]
                                       ([SiteName]
                                       ,[SiteURL]
                                       ,[ReportType]
                                       ,[Period]
                                       ,[IsAudit]
                                       ,[IsDownloaded]
                                       ,[DownloadedBy]
                                       ,[IsActive]
                                       ,[FileContent]
                                       ,[ReportUniuqeCode]
                                       ,[Created]
                                       ,[Modified])
                                 VALUES
                                       (@SiteName
                                       ,@SiteURL
                                       ,@ReportType
                                       ,@Period
                                       ,@IsAudit
                                       ,@IsDownloaded
                                       ,@DownloadedBy
                                       ,@IsActive
                                       ,@FileContent
                                       ,@ReportUniuqeCode
                                       ,@Created
                                       ,@Modified)";

                List<SqlParameter> parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@SiteName", reportTopInfo.SiteName));
                parameters.Add(new SqlParameter("@SiteURL", reportTopInfo.SiteUrl));
                parameters.Add(new SqlParameter("@ReportType", reportType));
                parameters.Add(new SqlParameter("@Period", period));
                parameters.Add(new SqlParameter("@IsAudit", isAudit));
                parameters.Add(new SqlParameter("@IsDownloaded", false));
                parameters.Add(new SqlParameter("@DownloadedBy", string.Empty));
                parameters.Add(new SqlParameter("@IsActive", isActive));
                parameters.Add(new SqlParameter("@FileContent", byteData));
                parameters.Add(new SqlParameter("@ReportUniuqeCode", reportUniqueCode));
                parameters.Add(new SqlParameter("@Created", DateTime.Now));
                parameters.Add(new SqlParameter("@Modified", DateTime.Now));
                rowsAffected = SqlHelper.ExecuteNonQuery(sConnectionString, CommandType.Text, strQuery, parameters.ToArray());
            }
            catch (Exception ex)
            {
                throw;
            }
            return rowsAffected;
        }

        /// <summary>
        /// Function to read the file record from SQL
        /// </summary>
        /// <param name="fileName">Name of the file along with the extension</param>
        /// <param name="siteName">Site Name</param>
        /// <param name="URL">URL of the site</param>
        /// <returns>Data table if the file is available, null if its not available</returns>
        public static DataTable ReadSecurableObjectFile(string fileName, string siteName, string URL)
        {
            DataTable dtable = null;
            try
            {
                string query = string.Format("SELECT  * FROM SecurableObjectsFiles WHERE SiteName = '{0}' AND FileName = '{1}' AND SiteURL = '{2}' AND IsActive = 1", siteName, fileName, URL);
                DataSet ds = SqlHelper.ExecuteDataset(sConnectionString, CommandType.Text, query);
                if (ds != null)
                    dtable = ds.Tables[0];
            }
            catch (Exception)
            {
                throw;
            }
            return dtable;
        }

        #endregion
    }
}
