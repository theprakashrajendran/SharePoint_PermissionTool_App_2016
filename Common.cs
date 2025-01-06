using Microsoft.SharePoint.Client;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;

namespace BASF.SharePoint.PermissionTool.Scheduler
{
    class Common
    {
        private static string sAppName = "";

        #region Check Document Library Exists
        public static bool CheckIfDocumentLibraryExists(ClientContext clientContext, string DocumentLibraryName)
        {
            try
            {
                ListCollection listCollection = clientContext.Web.Lists;
                clientContext.Load(listCollection, lists => lists.Include(list => list.Title).Where(list => list.Title == DocumentLibraryName));
                clientContext.ExecuteQuery();
                if (listCollection.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                SQL.InsertErrorLog("", sAppName, "Program.CheckIfDocumentLibraryExists", ex);
                return false;
            }
        }
        #endregion
        #region Create Document Library
        public static bool CreateDocumentLibrary(ClientContext clientContext, string DocumentLibraryName)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("Creating Document Library");

                Web web = clientContext.Web;
                ListCreationInformation creationInfo = new ListCreationInformation();
                creationInfo.Title = DocumentLibraryName;
                creationInfo.TemplateType = (int)ListTemplateType.DocumentLibrary;
                List list = web.Lists.Add(creationInfo);
                list.Update();
                clientContext.ExecuteQuery();

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Successfully Created Document Library");

                return true;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Failed to Create Document Library");
                SQL.InsertErrorLog("", sAppName, "Program.CreateDocumentLibrary", ex);
                return false;
            }
        }
        #endregion
        #region Update Document Library Permissions
        public static bool UpdateDocumentLibraryPermissions(ClientContext ctx, string DocumentLibraryName)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("Updating Document Library Permission");

                var web = ctx.Web;
                ctx.Load(ctx.Web, a => a.Lists, a => a.HasUniqueRoleAssignments);
                List list = ctx.Web.Lists.GetByTitle(DocumentLibraryName);
                ctx.Load(list, l => l.HasUniqueRoleAssignments);
                ctx.ExecuteQuery();
                //Stop Inheritance from parent
                if (!list.HasUniqueRoleAssignments)
                {
                    list.BreakRoleInheritance(false, false);
                    list.Update();
                    ctx.ExecuteQuery();
                }

                var roleAssignments = list.RoleAssignments;
                var user_group = web.AssociatedOwnerGroup;
                var roleDefBindCol = new RoleDefinitionBindingCollection(ctx);

                // Add Role Definition i.e Full Controls, Contribute or Read rights etc..
                roleDefBindCol.Add(web.RoleDefinitions.GetByType(RoleType.Administrator));
                roleAssignments.Add(user_group, roleDefBindCol);
                ctx.Load(roleAssignments);
                list.Update();
                ctx.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Failed to Update Document Library Permission");

                SQL.InsertErrorLog("", sAppName, "Program.UpdateDocumentLibraryPermissions", ex);
                return false;
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Successfully Updated Document Library Permission");
            return true;
        }
        #endregion

        #region Upload to SP Document Library
        public static bool UploadToSharePointDocumentLibrary(ClientContext context, string DocumentLibraryName, string FilePath)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("Uploading Excel File to Document Library");

                Web site = context.Web;
                string RootFolderRelativeUrl = DocumentLibraryName;
                Folder Folder = site.GetFolderByServerRelativeUrl(RootFolderRelativeUrl);
                FileCreationInformation newFile = new FileCreationInformation { Content = System.IO.File.ReadAllBytes(@FilePath), Url = Path.GetFileName(@FilePath), Overwrite = true };
                Folder.Files.Add(newFile);
                Folder.Update();
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Failed to Upload to Document Library");
                SQL.InsertErrorLog("", sAppName, "Program.UploadToSharePointDocumentLibrary", ex);
                return false;
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Successfully to Uploaded to Document Library");
            return true;
        }
        #endregion

        #region File Operations
        private static bool CheckAndUploadFile(string webUrl, string fileName)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Uploading Excel File");
                string DocumentLibraryName = "Security Reports"; //Store this value to resource file
                using (ClientContext _context = new ClientContext(webUrl))
                {

                    if (CheckIfDocumentLibraryExists(_context, DocumentLibraryName))
                    {
                        if (UpdateDocumentLibraryPermissions(_context, DocumentLibraryName))
                        {
                            UploadToSharePointDocumentLibrary(_context, DocumentLibraryName, fileName);
                        }
                    }
                    else
                    {
                        if (CreateDocumentLibrary(_context, DocumentLibraryName))
                        {
                            if (UpdateDocumentLibraryPermissions(_context, DocumentLibraryName))
                            {
                                UploadToSharePointDocumentLibrary(_context, DocumentLibraryName, fileName);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Failed to Upload Excel file");
                SQL.InsertErrorLog("", sAppName, "Program.CheckAndUploadFile", ex);
                return false;
            }
            return true;
        }
        #endregion

        #region Generate Excel File
        private static void GenerateExcel(DataTable dtSource, string reportType, ReportTopInfo reportTopdata, string dataLoadFrom, string period)
        {
            try
            {
                string reportUniqueCode = string.Empty;
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Generating Excel File");

                ExcelPackage pck = new ExcelPackage();

                string tableName = reportTopdata.Title;

                var ws = pck.Workbook.Worksheets.Add(tableName);
                ws.Cells[dataLoadFrom].LoadFromDataTable(dtSource, true);
                if (reportType == REPORTNAMES.reportNames.SECURABLEOBJECTREPORT.ToString())
                {
                    ws.Cells[1, 2].Value = reportTopdata.Title;
                    ws.Cells[2, 2].Value = "Site Collection URL:";
                    ws.Cells[2, 3].Value = reportTopdata.SiteUrl;
                    ws.Cells[3, 2].Value = "Owners:";
                    ws.Cells[3, 3].Value = reportTopdata.SiteOwners;
                    ws.Cells[2, 7].Value = "Date of creation:";
                    ws.Cells[2, 8].Value = reportTopdata.CreatedDate;
                    ws.Cells[3, 7].Value = "Storage quota/ usage GB/ usage %:";
                    ws.Cells[3, 8].Value = reportTopdata.Usage;

                    reportUniqueCode = Program.GenerateUniqueReportId(REPORTTYPE.reportType.SecurableObject.ToString(), period);
                }
                else if (reportType == REPORTNAMES.reportNames.PRINCIPALREPORT.ToString())
                {

                    ws.Cells[1, 2].Value = "Security Report: Principal Objects";
                    ws.Cells[2, 2].Value = "Site Collection URL:";
                    ws.Cells[2, 3].Value = reportTopdata.SiteUrl;
                    ws.Cells[3, 2].Value = "Owners:";
                    ws.Cells[3, 3].Value = reportTopdata.SiteOwners;
                    ws.Cells[2, 7].Value = "Date of creation:";
                    ws.Cells[2, 8].Value = reportTopdata.CreatedDate;
                    ws.Cells[3, 7].Value = "Storage quota/ usage GB/ usage %:";
                    ws.Cells[3, 8].Value = reportTopdata.Usage;

                    reportUniqueCode = Program.GenerateUniqueReportId(REPORTTYPE.reportType.PrincipalObject.ToString(), period);

                    /*ws.Cells[1, 2].Value = reportTopdata.Title;
                    ws.Cells[2, 2].Value = "Site Collection URL:";
                    ws.Cells[2, 3].Value = reportTopdata.SiteUrl;
                    ws.Cells[3, 2].Value = "Owners:";
                    ws.Cells[3, 3].Value = reportTopdata.SiteOwners;
                    ws.Cells[2, 7].Value = "Date of creation:";
                    ws.Cells[2, 8].Value = reportTopdata.CreatedDate;
                    ws.Cells[3, 7].Value = "Storage quota/ usage GB/ usage %:";
                    ws.Cells[3, 8].Value = reportTopdata.Usage;
                    ws.Cells[5, 3].Value = "Principal Information";
                    ws.Cells[5, 7].Value = "Securable Object Information";
                    ws.Cells[5, 11].Value = "Permission Information";*/
                }
                else if (reportType == REPORTNAMES.reportNames.STATISTICSANDUSAGEREPORT.ToString())
                {
                    ws.Cells[1, 2].Value = reportTopdata.Title;
                    ws.Cells[2, 2].Value = "Date of creation:";
                    ws.Cells[2, 3].Value = reportTopdata.CreatedDate;
                    ws.Cells[4, 2].Value = "Site Collection Information";
                    ws.Cells[4, 4].Value = "Site Owner Information";
                    ws.Cells[4, 10].Value = "Authorized users information";
                    ws.Cells[4, 10, 4, 13].Merge = true;
                    ws.Cells[4, 14].Value = "Usage information";
                }
                else if (reportType == REPORTNAMES.reportNames.ORPHANUSERREPORT.ToString())
                {
                    ws.Cells[1, 2].Value = reportTopdata.Title;
                    ws.Cells[2, 2].Value = "Date of creation:";
                    ws.Cells[2, 3].Value = reportTopdata.CreatedDate;
                    ws.Cells[2, 4].Value = "(does not include inactive users from within (nested) AD groups)";
                    ws.Cells[4, 3].Value = "User Information";
                    ws.Cells[4, 8].Value = "Site Collection Information";
                    ws.Cells[4, 10].Value = "Site Owner Information";
                }

                FormatExcelData(ws, reportType);

                StoreExcelInDB(pck, reportTopdata, reportType, period, true, true, reportUniqueCode);


                //string fileName = reportUniqueCode + "_" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss") + ".xlsx";
                //var folderPath = @"C:\Users\mariyafs\Downloads\BASF\" + fileName; //Need to modify to Application StartUp Path

                //try
                //{
                //    Stream stream = System.IO.File.Create(folderPath);
                //    pck.SaveAs(stream);
                //    stream.Close();
                //}
                //catch (Exception ex)
                //{
                //    SQL.InsertErrorLog("", sAppName, "Program.GenerateExcel.STREAM", ex);
                //    throw;
                //}




                //Console.ForegroundColor = ConsoleColor.Green;
                //Console.WriteLine("Excel Generated");

                //if (CheckAndUploadFile(reportTopdata.SiteUrl, folderPath))
                //{
                //    DeleteSecurityReportFile(folderPath);
                //}


            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(reportType +" Excel File Failed to Generate");
                SQL.InsertErrorLog("", sAppName, "Program.GenerateExcel", ex);
            }
        }
        #endregion

        #region Delete Security File
        private static void DeleteSecurityReportFile(string folderPath)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("Removing Dump files from Local Disk");
                if (System.IO.File.Exists(folderPath))
                    System.IO.File.Delete(folderPath);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Failed to Removing Dump files from Local Disk");
                SQL.InsertErrorLog("", sAppName, "Program.DeleteSecurityReportFile", ex);
                throw;
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Successfully Removed Dump files from Local Disk");
        }
        #endregion

        #region Format Excel Data
        private static void FormatExcelData(ExcelWorksheet ws, string reportType)
        {
            try
            {
                ws.Cells[6, 1, 6, 15].Style.Font.Bold = false;
                ws.Cells[6, 1, 6, 15].Style.Font.Size = 11;
                var cellStyle = ws.Cells.Style;

                cellStyle.VerticalAlignment = ExcelVerticalAlignment.Top;
                cellStyle.WrapText = false;
                cellStyle.ShrinkToFit = false;
                cellStyle.Font.Name = "Calibri";

                ws.Cells[1, 1, ws.Dimension.End.Row, ws.Dimension.End.Column].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                if (reportType == REPORTNAMES.reportNames.SECURABLEOBJECTREPORT.ToString())
                {
                    /*ws.Cells[6, 8, 6, 11].Style.TextRotation = 90;
                    ws.Cells[6, 8, 6, 11].Style.WrapText = true;
                    ws.Cells[6, 8, 6, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[6, 8, 6, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                    ws.Cells[6, 16, 6, 17].Style.TextRotation = 90;
                    ws.Cells[6, 16, 6, 17].Style.WrapText = true;
                    ws.Cells[6, 16, 6, 17].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[6, 16, 6, 17].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                    ws.Cells[5, 1, 5, 21].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells[5, 1, 5, 21].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                    ws.Cells[5, 1, ws.Dimension.End.Row, 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[5, 12, ws.Dimension.End.Row, 12].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[5, 22, ws.Dimension.End.Row, 22].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[5, 1, 5, 15].Style.Border.Bottom.Color.SetColor(Color.Black);
                    ws.Cells[5, 1, 5, 15].Style.Border.Top.Color.SetColor(Color.Black);

                    ws.Cells[5, 1, ws.Dimension.End.Row, 2].Style.Border.Left.Color.SetColor(Color.Black);
                    ws.Cells[5, 12, ws.Dimension.End.Row, 12].Style.Border.Left.Color.SetColor(Color.Black);
                    ws.Cells[5, 22, ws.Dimension.End.Row, 22].Style.Border.Left.Color.SetColor(Color.Black);*/

                    ws.Cells[1, 2].Style.Font.Bold = true;
                    ws.Cells[1, 2].Style.Font.Size = 14;
                    ws.Cells[2, 2].Style.Font.Bold = true;
                    ws.Cells[3, 2].Style.Font.Bold = true;
                    ws.Cells[2, 7].Style.Font.Bold = true;
                    ws.Cells[3, 7].Style.Font.Bold = true;

                    ws.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[3, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[2, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[6, 1, ws.Dimension.End.Row, 12].AutoFilter = true;
                    ws.Cells[6, 1, 6, 12].Style.Font.Bold = true;
                    //ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    ws.Cells.AutoFitColumns(3.71, 16.57);
                    /*ws.Column(8).Width = 6.57;
                    ws.Column(9).Width = 6.57;
                    ws.Column(10).Width = 6.57;
                    ws.Column(11).Width = 6.57;
                    ws.Column(16).Width = 3.71;
                    ws.Column(17).Width = 3.71;*/
                    ws.Cells[6, 13, 6, 13].Style.WrapText = true;
                    ws.Cells[6, 15, 6, 15].Style.WrapText = true;
                    ws.Cells[6, 18, 6, 21].Style.WrapText = true;
                }

                else if (reportType == REPORTNAMES.reportNames.PRINCIPALREPORT.ToString())
                {
                    ws.Cells[1, 2].Style.Font.Bold = true;
                    ws.Cells[1, 2].Style.Font.Size = 14;
                    ws.Cells[2, 2].Style.Font.Bold = true;
                    ws.Cells[3, 2].Style.Font.Bold = true;
                    ws.Cells[2, 7].Style.Font.Bold = true;
                    ws.Cells[3, 7].Style.Font.Bold = true;

                    ws.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[3, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[2, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[6, 1, ws.Dimension.End.Row, 12].AutoFilter = true;
                    ws.Cells[6, 1, 6, 12].Style.Font.Bold = true;
                    ws.Cells.AutoFitColumns(3.71, 16.57);
                    ws.Cells[6, 13, 6, 13].Style.WrapText = true;
                    ws.Cells[6, 15, 6, 15].Style.WrapText = true;
                    ws.Cells[6, 18, 6, 21].Style.WrapText = true;

                    /*ws.Cells[6, 6, 6, 6].Style.TextRotation = 90;
                    ws.Cells[6, 6, 6, 6].Style.WrapText = true;
                    ws.Cells[6, 6, 6, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[6, 6, 6, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    ws.Row(6).Height = 84.00;

                    ws.Cells[6, 10, 6, 10].Style.TextRotation = 90;
                    ws.Cells[6, 10, 6, 10].Style.WrapText = true;
                    ws.Cells[6, 10, 6, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[6, 10, 6, 10].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                    ws.Cells[5, 1, 5, 15].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells[5, 1, 5, 15].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                    ws.Cells[5, 1, ws.Dimension.End.Row, 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[5, 7, ws.Dimension.End.Row, 7].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[5, 11, ws.Dimension.End.Row, 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[5, 16, ws.Dimension.End.Row, 16].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //ws.Cells[6, 1, ws.Dimension.End.Row, 15].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    //ws.Cells[6, 1, ws.Dimension.End.Row, 15].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ws.Cells[5, 1, 5, 15].Style.Border.Bottom.Color.SetColor(Color.Black);
                    ws.Cells[5, 1, 5, 15].Style.Border.Top.Color.SetColor(Color.Black);

                    ws.Cells[5, 1, ws.Dimension.End.Row, 2].Style.Border.Left.Color.SetColor(Color.Black);
                    ws.Cells[5, 7, ws.Dimension.End.Row, 7].Style.Border.Left.Color.SetColor(Color.Black);
                    ws.Cells[5, 11, ws.Dimension.End.Row, 11].Style.Border.Left.Color.SetColor(Color.Black);
                    ws.Cells[5, 16, ws.Dimension.End.Row, 16].Style.Border.Left.Color.SetColor(Color.Black);
                    //ws.Cells[6, 1, ws.Dimension.End.Row, 15].Style.Border.Right.Color.SetColor(Color.Black);

                    ws.Cells[1, 2].Style.Font.Bold = true;
                    ws.Cells[1, 2].Style.Font.Size = 14;
                    ws.Cells[2, 2].Style.Font.Bold = true;
                    ws.Cells[3, 2].Style.Font.Bold = true;
                    ws.Cells[2, 7].Style.Font.Bold = true;
                    ws.Cells[3, 7].Style.Font.Bold = true;

                    ws.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[3, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[2, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[6, 1, ws.Dimension.End.Row, 15].AutoFilter = true;
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();*/
                }

                else if (reportType == REPORTNAMES.reportNames.STATISTICSANDUSAGEREPORT.ToString())
                {
                    ws.Cells[1, 2].Style.Font.Bold = true;
                    ws.Cells[1, 2].Style.Font.Size = 14;
                    ws.Cells[2, 2].Style.Font.Bold = true;
                    ws.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[5, 1, ws.Dimension.End.Row, ws.Dimension.End.Column - 2].AutoFilter = true;

                    ws.Cells[4, 1, 4, ws.Dimension.End.Column - 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells[4, 1, 4, ws.Dimension.End.Column - 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                    ws.Cells[4, 1, ws.Dimension.End.Row, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[4, 4, ws.Dimension.End.Row, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[4, 10, ws.Dimension.End.Row, 10].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[4, 14, ws.Dimension.End.Row, 14].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[4, 15, ws.Dimension.End.Row, 15].Style.Border.Left.Style = ExcelBorderStyle.Thin;

                    ws.Cells[4, 1, 4, ws.Dimension.End.Column - 1].Style.Border.Bottom.Color.SetColor(Color.Black);
                    ws.Cells[4, 1, 4, ws.Dimension.End.Column - 1].Style.Border.Top.Color.SetColor(Color.Black);

                    ws.Cells[5, 10, 5, 13].Style.TextRotation = 90;
                    ws.Cells[5, 10, 5, 13].Style.WrapText = true;
                    ws.Cells[5, 10, 5, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[5, 10, 5, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                    ws.Row(5).Height = 84.00;
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                }
                else if (reportType == REPORTNAMES.reportNames.ORPHANUSERREPORT.ToString())
                {
                    ws.Cells[1, 2].Style.Font.Bold = true;
                    ws.Cells[1, 2].Style.Font.Size = 14;
                    ws.Cells[2, 2].Style.Font.Bold = true;
                    ws.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells[2, 4].Style.Font.Italic = true;
                    ws.Cells[5, 10].Style.WrapText = true;
                    ws.Cells[5, 11].Style.WrapText = true;
                    ws.Cells[5, 1, ws.Dimension.End.Row, 11].AutoFilter = true;

                    ws.Cells[4, 1, 4, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells[4, 1, 4, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                    ws.Cells[4, 1, ws.Dimension.End.Row, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[4, 8, ws.Dimension.End.Row, 8].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[4, 10, ws.Dimension.End.Row, 10].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[4, 12, ws.Dimension.End.Row, 12].Style.Border.Left.Style = ExcelBorderStyle.Thin;

                    ws.Cells[4, 1, 4, 11].Style.Border.Bottom.Color.SetColor(Color.Black);
                    ws.Cells[4, 1, 4, 11].Style.Border.Top.Color.SetColor(Color.Black);

                    ws.Cells[5, 5, 5, 7].Style.TextRotation = 90;
                    ws.Cells[5, 5, 5, 7].Style.WrapText = true;
                    ws.Cells[5, 5, 5, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.Cells[5, 5, 5, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                    ws.Row(5).Height = 84.00;
                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                }

                ws.Column(5).BestFit = true;
                ws.Column(6).BestFit = true;
                ws.Column(7).BestFit = true;
                ws.Column(8).BestFit = true;
                ws.Column(9).BestFit = true;
                ws.Column(10).BestFit = true;
                ws.Column(11).BestFit = true;
                ws.Column(12).BestFit = true;
                ws.Column(13).BestFit = true;
                ws.Column(14).BestFit = true;
                ws.Column(15).BestFit = true;


            }
            catch (Exception ex)
            {
                SQL.InsertErrorLog("", sAppName, "Program.FormatExcelData", ex);
            }
        }
        #endregion

        #region Export To Excel
        public static bool ExportToExcel(DataTable dtRes, string reportType, ReportTopInfo reportTopdata, string period)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.DarkMagenta;
                Console.WriteLine("Exporting Excel file");

                string fileName = string.Empty;
                string dataToLoadFrom = string.Empty;

                #region SECURE OBJECT REPORT
                if (reportType == REPORTNAMES.reportNames.SECURABLEOBJECTREPORT.ToString())
                {
                    fileName = "Securable Objects";
                    dataToLoadFrom = "A6";

                    dtRes.Columns["ParentWebUrl"].ColumnName = "Parent Web Url";
                    dtRes.Columns["ParentName"].ColumnName = "Parent Name";
                    dtRes.Columns["ObjectUrl"].ColumnName = "Object Url";
                    dtRes.Columns["ObjectName"].ColumnName = "Object Name";
                    dtRes.Columns["ObjectType"].ColumnName = "Object Type";
                    dtRes.Columns["PrincipalName"].ColumnName = "Principal Name";
                    dtRes.Columns["PrincipalEmail"].ColumnName = "Principal Email";
                    dtRes.Columns["PrincipalAlias"].ColumnName = "Principal Alias";
                    dtRes.Columns["PrincipalType"].ColumnName = "Principal Type";
                    dtRes.Columns["PrincipalSource"].ColumnName = "Principal Source";
                    dtRes.Columns["PermissionLevel"].ColumnName = "Permission Level";
                    dtRes.Columns["MemberCount"].ColumnName = "Member Count";
                }
                #endregion

                #region PRINCIPAL REPORT
                else if (reportType == REPORTNAMES.reportNames.PRINCIPALREPORT.ToString())
                {
                    fileName = "Principals";
                    dataToLoadFrom = "A6";


                    dtRes = dtRes.DefaultView.ToTable(true, "SID", "RowType",
                                       "PrincipalName", "PrincipalAlias", "PrincipalType", "PrincipalCategory",
                                       "ObjectId", "ObjectName", "ObjectType", "HasChildUniquePermissions", "PermissionLevel",
                                       "AssignedThroughSPGroup", "IsNestedADGroup", "ParentADGroup", "ItemLink");



                    dtRes.Columns["SID"].SetOrdinal(0);
                    dtRes.Columns["RowType"].SetOrdinal(1);
                    dtRes.Columns["PrincipalName"].SetOrdinal(2);
                    dtRes.Columns["PrincipalAlias"].SetOrdinal(3);
                    dtRes.Columns["PrincipalType"].SetOrdinal(4);
                    dtRes.Columns["PrincipalCategory"].SetOrdinal(5);
                    dtRes.Columns["ObjectId"].SetOrdinal(6);
                    dtRes.Columns["ObjectName"].SetOrdinal(7);
                    dtRes.Columns["ObjectType"].SetOrdinal(8);
                    dtRes.Columns["HasChildUniquePermissions"].SetOrdinal(9);
                    dtRes.Columns["PermissionLevel"].SetOrdinal(10);
                    dtRes.Columns["AssignedThroughSPGroup"].SetOrdinal(11);
                    dtRes.Columns["IsNestedADGroup"].SetOrdinal(12);
                    dtRes.Columns["ParentADGroup"].SetOrdinal(13);
                    dtRes.Columns["ItemLink"].SetOrdinal(14);



                    dtRes.Columns["SID"].ColumnName = "Note";
                    dtRes.Columns["RowType"].ColumnName = "Row Type";
                    dtRes.Columns["PrincipalName"].ColumnName = "Principal Name";
                    dtRes.Columns["PrincipalAlias"].ColumnName = "Principal Alias";
                    dtRes.Columns["PrincipalType"].ColumnName = "Principal Type";
                    dtRes.Columns["PrincipalCategory"].ColumnName = "Is internal, contractor or external";
                    dtRes.Columns["ObjectId"].ColumnName = "Object Id";
                    dtRes.Columns["ObjectName"].ColumnName = "Object Name";
                    dtRes.Columns["ObjectType"].ColumnName = "Object Type";
                    dtRes.Columns["HasChildUniquePermissions"].ColumnName = "Has children  with unique permissions";
                    dtRes.Columns["PermissionLevel"].ColumnName = "Permission Level";
                    dtRes.Columns["AssignedThroughSPGroup"].ColumnName = "Assigned through SharePoint Group";
                    dtRes.Columns["IsNestedADGroup"].ColumnName = "Is nested AD Group";
                    dtRes.Columns["ParentADGroup"].ColumnName = "Parent AD Group";
                    dtRes.Columns["ItemLink"].ColumnName = "Link to item";
                }

                #endregion

                #region STATISTICS AND USAGE REPORT
                else if (reportType == REPORTNAMES.reportNames.STATISTICSANDUSAGEREPORT.ToString())
                {
                    fileName = "Statistics and Usage";
                    dataToLoadFrom = "A5";

                    dtRes = dtRes.DefaultView.ToTable(true, "SID", "Id",
                                       "URL", "PrimaryOwnerAlias", "SecondaryOwnerAlias", "FullControlUsers",
                                       "Quota", "StorageGB", "StoragePercentage", "InternalUsers", "Contractors",
                                       "Externals", "NestedADGroupUsers", "LastWriteAccess");



                    dtRes.Columns["SID"].SetOrdinal(0);
                    dtRes.Columns["Id"].SetOrdinal(1);
                    dtRes.Columns["URL"].SetOrdinal(2);
                    dtRes.Columns["PrimaryOwnerAlias"].SetOrdinal(3);
                    dtRes.Columns["SecondaryOwnerAlias"].SetOrdinal(4);
                    dtRes.Columns["FullControlUsers"].SetOrdinal(5);
                    dtRes.Columns["Quota"].SetOrdinal(6);
                    dtRes.Columns["StorageGB"].SetOrdinal(7);
                    dtRes.Columns["StoragePercentage"].SetOrdinal(8);
                    dtRes.Columns["InternalUsers"].SetOrdinal(9);
                    dtRes.Columns["Contractors"].SetOrdinal(10);
                    dtRes.Columns["Externals"].SetOrdinal(11);
                    dtRes.Columns["NestedADGroupUsers"].SetOrdinal(12);
                    dtRes.Columns["LastWriteAccess"].SetOrdinal(13);



                    dtRes.Columns["SID"].ColumnName = "Note";
                    dtRes.Columns["Id"].ColumnName = "Id";
                    dtRes.Columns["URL"].ColumnName = "URL";
                    dtRes.Columns["PrimaryOwnerAlias"].ColumnName = "Primary Owner Alias";
                    dtRes.Columns["SecondaryOwnerAlias"].ColumnName = "Secondary Owner Alias";
                    dtRes.Columns["FullControlUsers"].ColumnName = "Addtional users with full control";
                    dtRes.Columns["Quota"].ColumnName = "Quota";
                    dtRes.Columns["StorageGB"].ColumnName = "Storage GB";
                    dtRes.Columns["StoragePercentage"].ColumnName = "Storage %";
                    dtRes.Columns["InternalUsers"].ColumnName = "Internal users";
                    dtRes.Columns["Contractors"].ColumnName = "Contractors";
                    dtRes.Columns["Externals"].ColumnName = "Externals";
                    dtRes.Columns["NestedADGroupUsers"].ColumnName = "Users in (nested) AD groups";
                    dtRes.Columns["LastWriteAccess"].ColumnName = "Last write access";
                }

                #endregion

                #region ORPHAN USER REPORT
                else if (reportType == REPORTNAMES.reportNames.ORPHANUSERREPORT.ToString())
                {

                    fileName = "Orphaned users";
                    dataToLoadFrom = "A5";

                    //Note == SID
                    //Is internal, contractor or external	Is expired == PrincipalCategory
                    dtRes = dtRes.DefaultView.ToTable(true, "SID", "RowType",
                                       "PrincipalName", "PrincipalAlias", "PrincipalCategory", "IsExpired",
                                       "IsDisabled", "Name", "URL", "PrimaryOwnerAlias", "SecondaryOwnerAlias");



                    dtRes.Columns["SID"].SetOrdinal(0);
                    dtRes.Columns["RowType"].SetOrdinal(1);
                    dtRes.Columns["PrincipalName"].SetOrdinal(2);
                    dtRes.Columns["PrincipalAlias"].SetOrdinal(3);
                    dtRes.Columns["PrincipalCategory"].SetOrdinal(4);
                    dtRes.Columns["IsExpired"].SetOrdinal(5);
                    dtRes.Columns["IsDisabled"].SetOrdinal(6);
                    dtRes.Columns["Name"].SetOrdinal(7);
                    dtRes.Columns["URL"].SetOrdinal(8);
                    dtRes.Columns["PrimaryOwnerAlias"].SetOrdinal(9);
                    dtRes.Columns["SecondaryOwnerAlias"].SetOrdinal(10);

                    dtRes.Columns["SID"].ColumnName = "Note";
                    dtRes.Columns["RowType"].ColumnName = "Row Type";
                    dtRes.Columns["PrincipalName"].ColumnName = "Principal Name";
                    dtRes.Columns["PrincipalAlias"].ColumnName = "Principal Alias";
                    dtRes.Columns["PrincipalCategory"].ColumnName = "Is internal, contractor or external";
                    dtRes.Columns["IsExpired"].ColumnName = "Is expired";
                    dtRes.Columns["IsDisabled"].ColumnName = "Is disabled";
                    dtRes.Columns["Name"].ColumnName = "Name";
                    dtRes.Columns["URL"].ColumnName = "URL";
                    dtRes.Columns["PrimaryOwnerAlias"].ColumnName = "Primary Owner Alias";
                    dtRes.Columns["SecondaryOwnerAlias"].ColumnName = "Secondary Owner Alias";
                }

                #endregion

                GenerateExcel(dtRes, reportType, reportTopdata, dataToLoadFrom, period);

                dtRes.DefaultView.Sort = "Principal Name ASC";
                dtRes = dtRes.DefaultView.ToTable();

                dtRes.Columns["Principal Name"].SetOrdinal(0);
                dtRes.Columns["Principal Email"].SetOrdinal(1);
                dtRes.Columns["Principal Alias"].SetOrdinal(2);
                dtRes.Columns["Principal Type"].SetOrdinal(3);
                dtRes.Columns["Principal Source"].SetOrdinal(4);
                        

                GenerateExcel(dtRes, REPORTNAMES.reportNames.PRINCIPALREPORT.ToString(), reportTopdata, dataToLoadFrom, period);

            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Excel File Exporting Failed");
                SQL.InsertErrorLog("", sAppName, "Program.ExportToExcel", ex);
                Console.WriteLine(ex.Message, ex);
                return false;
            }
            // Console.ForegroundColor = ConsoleColor.Green;
            //   Console.WriteLine("Excel File Exported Successfully");
            return true;
        }
        #endregion

        #region Store Excel in DB

        /// <summary>
        /// Function to save excel data as BLOB into SQL
        /// </summary>
        /// <param name="pck">Excel Package</param>
        /// <param name="fileName">Name of the file</param>
        /// <param name="reportTopInfo">Report information</param>
        /// <param name="reportType">Report Type</param>
        /// <param name="period">Period of report</param>
        /// <param name="isAudit">Is the report is for Audit?</param>
        /// <param name="isActive">Is the report is Active?</param>
        /// <param name="reportUniqueCode">Report unique code</param>
        /// <returns>1 if the file is saved, -1 if the file is not saved</returns>
        private static int StoreExcelInDB(ExcelPackage pck, ReportTopInfo reportTopInfo, string reportType, string period, bool isAudit, bool isActive, string reportUniqueCode)
        {
            try
            {
                using (BinaryReader br = new BinaryReader(pck.Stream))
                {
                    byte[] byteData = pck.GetAsByteArray();

                    // Writting file to DB
                    return SQL.InsertSecurableObjectFile(byteData, reportTopInfo, reportType, period, isAudit, isActive, reportUniqueCode);
                }
            }
            catch (Exception)
            {
                return -1;
            }
        }

        #endregion

        #region Save Excel from DB

        private static void SaveExcelFromDB(string fileName, string siteName, string siteURL, string filePath)
        {
            // Reading file from DB
            DataTable dtFileTable = SQL.ReadSecurableObjectFile(fileName, siteName, siteURL);
            if (dtFileTable != null && dtFileTable.Rows.Count > 0)
            {
                byte[] byteData = dtFileTable.Rows[0]["FileContent"] as byte[];
                if (byteData != null)
                {
                    filePath = filePath.EndsWith(@"\") ? filePath : filePath + @"\";
                    using (FileStream fs = new FileStream(filePath + fileName, FileMode.CreateNew, FileAccess.Write))
                    {
                        fs.Write(byteData, 0, byteData.Length);
                        fs.Close();
                    }
                }
            }
        }

        #endregion

        #region Object Collection to DataTable

        public static DataTable ConvertToDataTable<T>(IList<T> list)
        {
            DataTable table = CreateTable<T>();
            Type entityType = typeof(T);
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(entityType);

            foreach (T item in list)
            {
                DataRow row = table.NewRow();

                foreach (PropertyDescriptor prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item);
                }

                table.Rows.Add(row);
            }

            return table;
        }

        public static DataTable CreateTable<T>()
        {
            Type entityType = typeof(T);
            DataTable table = new DataTable(entityType.Name);
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(entityType);

            foreach (PropertyDescriptor prop in properties)
            {
                // HERE IS WHERE THE ERROR IS THROWN FOR NULLABLE TYPES
                table.Columns.Add(prop.Name, prop.PropertyType);
            }

            return table;
        }

        #endregion
    }

    #region Getting the first/last day of the week or month with DateTime
    public static partial class DateTimeExtensions
    {
        public static DateTime FirstDayOfWeek(this DateTime dt)
        {
            var culture = System.Threading.Thread.CurrentThread.CurrentCulture;
            var diff = dt.DayOfWeek - culture.DateTimeFormat.FirstDayOfWeek;
            if (diff < 0)
                diff += 7;
            return dt.AddDays(-diff).Date;
        }

        public static DateTime LastDayOfWeek(this DateTime dt)
        {
            return dt.FirstDayOfWeek().AddDays(6);
        }

        public static DateTime FirstDayOfMonth(this DateTime dt)
        {
            return new DateTime(dt.Year, dt.Month, 1);
        }

        public static DateTime LastDayOfMonth(this DateTime dt)
        {
            return dt.FirstDayOfMonth().AddMonths(1).AddDays(-1);
        }

        public static DateTime FirstDayOfNextMonth(this DateTime dt)
        {
            return dt.FirstDayOfMonth().AddMonths(1);
        }
    }
    #endregion
}
