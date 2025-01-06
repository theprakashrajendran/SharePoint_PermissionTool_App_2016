using System.Text;
using System;


namespace BASF.SharePoint.PermissionTool.Scheduler
{
    public static class ErrorHandling
    {
        #region LOGGING

        #region ULS LOG
        /// <summary>
        /// Will log the messages in the Sharepoint LOGS folder in the Sharepoint's log files
        /// </summary>
        /// <param name="text">message</param>
        public static void WriteLog(string sourcemethod, Exception ex)
        {
             BuildMessage(ex);
        }
        #endregion

        #region LIST LOGGING
        /// <summary>
        /// Will log the messages in the Sharepoint LOGS folder in the Sharepoint's log files
        /// </summary>
        /// <param name="text">message</param>
        /// <param name="ex">exception object - Pass NULL if nothing is present</param>
        /// <param name="aspireWeb">SPWeb object</param>
        //public static void WriteLog(string text, Exception ex, SPWeb aspireWeb)
        //{
        //    //try
        //    //{
        //    //    SPSecurity.RunWithElevatedPrivileges(delegate()
        //    //    {
        //    //        using (SPSite site = new SPSite(aspireWeb.Url))
        //    //        {
        //    //            using (SPWeb elevatedWeb = site.OpenWeb())
        //    //            {
        //    //                SPList lstOnlineServices = elevatedWeb.Lists.TryGetList("ExceptionLogs");
        //    //                if (lstOnlineServices != null)
        //    //                {
        //    //                    SPListItem item = lstOnlineServices.Items.Add();

        //    //                    if (text.Length > 240)
        //    //                    {
        //    //                        item["Title"] = text.Substring(0, 239);
        //    //                    }
        //    //                    else
        //    //                    {
        //    //                        item["Title"] = text;
        //    //                    }
        //    //                    if (ex != null)
        //    //                    {
        //    //                        item["StackTrace"] = Convert.ToString(ex.StackTrace);
        //    //                        item["Message"] = BuildMessage(ex);
        //    //                    }
        //    //                    item["Username"] = aspireWeb.CurrentUser;

        //    //                    site.AllowUnsafeUpdates = true;
        //    //                    elevatedWeb.AllowUnsafeUpdates = true;
        //    //                    //SPUtility.ValidateFormDigest();                            
        //    //                    item.Update();
        //    //                    elevatedWeb.AllowUnsafeUpdates = false;
        //    //                    site.AllowUnsafeUpdates = false;
        //    //                }
        //    //            }
        //    //        }
        //    //    });
        //    //}
        //    //catch (Exception ex1)
        //    //{
        //    //    WriteLog("ERROR::WriteLog::Aspire MySite List Insert Error::" + ex1.Message);
        //    //    WriteLog("ERROR::WriteLog::Aspire MySite List::Main Mesage" + ex.Message);
        //    //}
        //}
        #endregion

        #region BUILD EXCEPTION
        /// <summary>
        /// Build exception in a format
        /// </summary>
        /// <param name="objChainedException">exception</param>
        /// <returns></returns>
        public static string BuildMessage(Exception objChainedException)
        {
            StringBuilder objChainedMessage = new StringBuilder();
            int objChainedMessageNum = 1;
            System.Exception inner = null;
            inner = objChainedException;
            while (inner != null)
            {
                objChainedMessage.Append(inner.GetType().ToString() + "" + Environment.NewLine + "").Append(objChainedMessageNum).Append(") ").Append(inner.Message);
                objChainedMessage.Append("" + Environment.NewLine + "");
                inner = inner.InnerException;
                objChainedMessageNum += 1;
                objChainedMessage.Append("" + Environment.NewLine + "----------------------" + Environment.NewLine + "");
            }

            if (objChainedException.StackTrace != null && objChainedException.StackTrace.Length > 0)
            {
                objChainedMessage.Append(objChainedException.StackTrace);
            }
            else
            {
                objChainedMessage.Append(objChainedException.InnerException.StackTrace);
            }
            return objChainedMessage.ToString();
        }
        #endregion

        #endregion
    }
}
