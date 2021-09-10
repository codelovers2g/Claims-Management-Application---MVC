using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using Buisness.Services;
using Common.Case.Models;
using Common.Case.Template;
using Common.Master;
using Common.QuickBooks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;

namespace SchedulingApp.CommonClass
{
    public class OpenXmlMailMerge
    {
        private static OpenXmlMailMerge instance;

        private static object lockPad = new object();
        private OpenXmlMailMerge()
        {
        }


        public static OpenXmlMailMerge GetOpenXmlMailMerge
        {
            get
            {
                lock (lockPad)
                {
                    if (instance == null)
                    {
                        instance = new OpenXmlMailMerge();
                    }

                    return instance;
                }
            }
        }

        public void copyfile(string pWordDoc, string TargetpWordDoc)
        {
            if (System.IO.File.Exists(pWordDoc))
            {
                System.IO.File.Copy(pWordDoc, TargetpWordDoc);
            }
            //================replace merged field for template==================================//
            using (WordprocessingDocument document1 = WordprocessingDocument.Open(TargetpWordDoc, true))
            {
                document1.MainDocumentPart.Document.Save();
            }
        }
        public void copyExcelfile(string pWordDoc, string TargetpWordDoc)
        {
            // Create excel file on physical disk  
            File.Copy(pWordDoc, TargetpWordDoc);
            //FileStream objFileStrm = File.Copy(pWordDoc, TargetpWordDoc);
            //System.IO.File.Copy(pWordDoc, TargetpWordDoc);
            //objFileStrm.Close();
        }
        public bool OpenXmlMergeForUserReport(string pWordDoc, string TargetpWordDoc, List<IMEDataReport2Updated> _IMEDataReport2, IMEReportFiltter IMEReportFiltter)
        {
            bool _return = false;
            try
            {
                if (System.IO.File.Exists(pWordDoc))
                {
                    System.IO.File.Copy(pWordDoc, TargetpWordDoc);
                }
                using (WordprocessingDocument document1 = WordprocessingDocument.Open(TargetpWordDoc, true))
                {
                    document1.FillMergeFieldsForUserReport(_IMEDataReport2, IMEReportFiltter);
                    document1.MainDocumentPart.Document.Save();
                }
                return true;
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in OpenXMlMailMerge/OpenXmlMergeForUserReport Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message);

                _return = false;
            }
            return _return;
        }
        public bool OpenXmlMerge(string pWordDoc, string TargetpWordDoc, Dictionary<string, string> pDictionaryMerge, NYCTemplate nYCTemplateObj, Dictionary<string, string> UniversalDictionaryMerge)
        {
            bool _return = false;
            try
            {
                //if (System.IO.File.Exists(pWordDoc))
                //{
                //    System.IO.File.Copy(pWordDoc, TargetpWordDoc);
                //}
                //================replace merged field for template==================================//
                using (WordprocessingDocument document1 = WordprocessingDocument.Open(TargetpWordDoc, true))
                {
                    document1.FillMergeFields(pDictionaryMerge, nYCTemplateObj);
                    document1.MainDocumentPart.Document.Save();
                }
                // problem template mergefield is not working because these are simole text not merged field so this funcation to replace the text
                using (WordprocessingDocument document1 = WordprocessingDocument.Open(TargetpWordDoc, true))
                {
                    string docText = null;
                    using (StreamReader sr = new StreamReader(document1.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd();
                    }
                    CustomLogger.LogError.GetLogger.Log(null, "Before orignal text :" + docText);

                    StringBuilder _log = new StringBuilder();
                    foreach (KeyValuePair<string, string> entry in UniversalDictionaryMerge)
                    {
                        string pattern = String.Format("(?<!\\S){0}(?!\\S)", entry.Key);
                        Regex regexText = new Regex(entry.Key, RegexOptions.IgnoreCase);
                        
                        if (regexText.IsMatch(docText))
                        {
                            string value = "";
                            if (entry.Value != null)
                            {
                                value = entry.Value.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("\'", "&apos;").Replace("<", "&lt;").Replace(">", "&gt;").ToString();
                                if (entry.Key == "Claimant.MultilineAddress" || entry.Key == "Attorney.MultilineAddress" || entry.Key == "Client.MultilineAddress" || entry.Key == "ServiceProvider.MultilineAddress")
                                {
                                    value = value.Replace("^^", "<w:br/>");
                                }
                                if (entry.Key == "Charges.Activity" || entry.Key == "Client.Charges")
                                {
                                    value = value.Replace("^^", "<w:br/>");
                                }
                                if (entry.Key == "ServiceProvider.MailingAddress")
                                {
                                    value = value.Replace("^^", "<w:br/>");
                                }
                            }
                            _log.Append(string.Format("Mtach Key {0} replace with {1} {2}", entry.Key, entry.Value == null ? "" : entry.Value, Environment.NewLine));
                            docText = regexText.Replace(docText, value);
                        }
                        else
                        {
                            //  _log.Append(string.Format("Not Mtach Key {0} replace with {1} {2}", entry.Key, entry.Value == null ? "" : entry.Value, Environment.NewLine));
                        }

                    }
                    CustomLogger.LogError.GetLogger.Log(null, _log.ToString());
                    CustomLogger.LogError.GetLogger.Log(null, "After orignal text :" + docText);

                    using (StreamWriter sw = new StreamWriter(document1.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {

                CustomLogger.LogError.GetLogger.Log(ex, "Error in OpenXmlMerge Function");
                try
                {
                    int _index = 0;
                    if (ex != null)
                    {
                        CustomLogger.LogError.GetLogger.Log(ex, "Error in OpenXmlMerge Function.");
                        QuickBookLogCQB _QuickBookLogCQB = new QuickBookLogCQB();
                        _QuickBookLogCQB.AccessToken = " "; //acces_token variable is defined in AppController as this controller is inherited from AppController
                        _QuickBookLogCQB.SyncRequest = "12341";
                        _QuickBookLogCQB.RealmId = " ";
                        _QuickBookLogCQB.CreatedDate = DateTime.UtcNow;
                        _QuickBookLogCQB.SyncUptoDate = null;
                        _QuickBookLogCQB.FailureReason = "Inner, Reason: " + Convert.ToString(ex?.InnerException?.Message == null ? ex?.Message : ex.InnerException?.Message);

                        QuickBooksManager qb = new QuickBooksManager();
                        bool SyncDetails = qb.QuickBooksSyncLog(_QuickBookLogCQB);
                    }
                }
                catch (Exception ex1)
                {
                    CustomLogger.LogError.GetLogger.Log(ex1, "Error in openXMLMailMerge/OpenXmlMerge writing error Function");
                    CustomLogger.LogError.GetLogger.Log(ex1, ex1.InnerException.Message.ToString());                   
                }

                _return = false;
            }
            return _return;
        }
        public bool OpenXmlMergeForReports(string pWordDoc, string TargetpWordDoc, List<IMEDataReport2Updated> _IMEDataReport2, IMEReportFiltter IMEReportFiltter)
        {
            bool _return = false;
            try
            {
                if (System.IO.File.Exists(pWordDoc))
                {
                    System.IO.File.Copy(pWordDoc, TargetpWordDoc);
                }
                using (WordprocessingDocument document1 = WordprocessingDocument.Open(TargetpWordDoc, true))
                {
                    document1.FillMergeFieldsForReports(_IMEDataReport2, IMEReportFiltter);
                    document1.MainDocumentPart.Document.Save();
                }
                return true;
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in openXMLMailMerge/OpenXmlMergeForReports writing error Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                _return = false;
            }
            return _return;
        }
        public bool OpenXmlMergeForExportReport(string pWordDoc, string TargetpWordDoc, List<CaseReport> _IMEDataReport2, IMEReportFiltter IMEReportFiltter)
        {
            bool _return = false;
            try
            {
                if (System.IO.File.Exists(pWordDoc))
                {
                    System.IO.File.Copy(pWordDoc, TargetpWordDoc);
                }
                using (WordprocessingDocument document1 = WordprocessingDocument.Open(TargetpWordDoc, true))
                {
                    document1.FillMergeFieldsForExportReport(_IMEDataReport2, IMEReportFiltter);
                    document1.MainDocumentPart.Document.Save();
                }
                return true;
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in openXMLMailMerge/OpenXmlMergeForExportReport writing error Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                _return = false;
            }
            return _return;
        }
        public bool OpenXmlMergeForIMEReportCSV(string pWordDoc, string TargetpWordDoc, List<IMEDataReport2> _IMEDataReport2, IMEReportFiltter IMEReportFiltter)
        {
            bool _return = false;
            try
            {
                if (System.IO.File.Exists(pWordDoc))
                {
                    System.IO.File.Copy(pWordDoc, TargetpWordDoc);
                }
                IMEDataReportCSV(TargetpWordDoc, _IMEDataReport2, IMEReportFiltter);
                return true;
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in openXMLMailMerge/OpenXmlMergeForIMEReportCSV writing error Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                _return = false;
            }
            return _return;
        }
        public void IMEDataReportCSV(string path, List<IMEDataReport2> _IMEDataReport2, IMEReportFiltter IMEReportFiltter)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(path);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                // get number of rows in the sheet
                int rows = worksheet.Dimension.Rows; // 10
                int cells = worksheet.Dimension.Columns; // 10

                // loop through the worksheet rows
                int i = 1;
                foreach (var item in _IMEDataReport2)
                {
                    i++;
                    worksheet.Cells[i, 1].Value = item.WCB;
                    worksheet.Cells[i, 2].Value = item.Claim;
                    worksheet.Cells[i, 3].Value = item.CaseNo;
                    worksheet.Cells[i, 4].Value = item.IMEEntityName;
                    worksheet.Cells[i, 5].Value = item.RegistrationNumber;
                    worksheet.Cells[i, 6].Value = item.CarrierCaseNumber;
                    worksheet.Cells[i, 7].Value = item.ClaimantFullName;
                    worksheet.Cells[i, 8].Value = item.DoctorName;
                    worksheet.Cells[i, 9].Value = item.DoctorWCB;
                    worksheet.Cells[i, 10].Value = item.TypeOfService;
                    worksheet.Cells[i, 11].Value = item.CaseStatus;
                    worksheet.Cells[i, 12].Value = item.DoctorSpecialty;
                    worksheet.Cells[i, 13].Value = item.DateofService;
                    worksheet.Cells[i, 14].Value = item.TimeofService;
                    worksheet.Cells[i, 15].Value = item.RequestingEntity;
                    worksheet.Cells[i, 16].Value = item.DateReferral;
                    worksheet.Cells[i, 17].Value = item.Dateofnotice;
                    worksheet.Cells[i, 18].Value = item.IME4signed;
                    worksheet.Cells[i, 19].Value = item.IME4sent;
                    worksheet.Cells[i, 20].Value = item.Amountbilled;
                    worksheet.Cells[i, 21].Value = item.ProviderPayment;
                    worksheet.Cells[i, 22].Value = item.IME4signedPOI;
                    worksheet.Cells[i, 23].Value = item.Starttime;
                    worksheet.Cells[i, 24].Value = item.Endtime;
                    worksheet.Cells[i, 25].Value = item.TotalTimeSpent;
                    worksheet.Cells[i, 26].Value = item.ApptStatus;
                    worksheet.Cells[i, 27].Value = item.streetaddress;
                    worksheet.Cells[i, 28].Value = item.citystatezipcode;
                }
                // save changes
                package.Save();
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in openXMLMailMerge/IMEDataReportCSV writing error Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
            }
        }
        public bool ReplaceInExcel(string path, Dictionary<string, string> pDictionaryMerge, NYCTemplate nYCTemplateObj, Dictionary<string, string> UniversalDictionaryMerge)
        {
            bool _return = true;
            try
            {
                FileInfo fileInfo = new FileInfo(path);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                // get number of rows in the sheet
                int rows = worksheet.Dimension.Rows; // 10
                int cells = worksheet.Dimension.Columns; // 10

                // loop through the worksheet rows
                for (int i = 1; i <= rows; i++)
                {
                    // replace occurences
                    for (int j = 1; j <= cells; j++)
                    {
                        var check = worksheet.Cells[i, j].Value?.ToString();
                        foreach (var item in pDictionaryMerge)
                        {
                            if (check != null)
                            {

                                if ((worksheet.Cells[i, j].Value != null) && (worksheet.Cells[i, j].Value.ToString().Contains(item.Key)) && (item.Key == "TableStart:HealthCarePractitioners"))
                                {
                                    //worksheet.Cells[i, j].Value = "";
                                    if (nYCTemplateObj.HealthCarePractioners != null)
                                    {
                                        worksheet.InsertRow(i, nYCTemplateObj.HealthCarePractioners.Count - 1);
                                        rows = worksheet.Dimension.Rows;
                                        foreach (var healthCare in nYCTemplateObj.HealthCarePractioners)
                                        {
                                            int k = 1;
                                            worksheet.Cells[i, k++].Value = "Dr. ";
                                            worksheet.Cells[i, k++].Value = healthCare.ServiceProviderFirstName;
                                            worksheet.Cells[i, k++].Value = healthCare.ServiceProviderLastName;
                                            string[] cityState = healthCare.ServiceProviderAddress.Split('~');
                                            worksheet.Cells[i, k++].Value = cityState[0];
                                            worksheet.Cells[i, k++].Value = "";
                                            worksheet.Cells[i, k++].Value = cityState[1];
                                            i++;
                                        }
                                        i--;
                                    }
                                    else
                                    {
                                        worksheet.Cells[i, j].Value = "";
                                    }

                                }
                                else if (worksheet.Cells[i, j].Value != null && worksheet.Cells[i, j].Value.ToString().Contains(item.Key))
                                {
                                    check = check.Replace(item.Key, item.Value);
                                    worksheet.Cells[i, j].Value = check;
                                }
                            }
                        }
                    }
                }

                // save changes
                package.Save();
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in openXMLMailMerge/ReplaceInExcel writing error Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                _return = false;
            }
            return _return;
        }

        public bool AppointmentLogExcelReport(string TempPath, string FilePath, List<IMEDataReport2Updated> _IMEDataReport2, IMEReportFiltter IMEReportFiltter)
        {
            bool _return = true;
            try
            {
                FileInfo fileInfo = new FileInfo(TempPath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                // get number of rows in the sheet
                int rows = worksheet.Dimension.Rows; // 10
                int cells = worksheet.Dimension.Columns; // 10
                int i = 3;

                foreach (var item in _IMEDataReport2)
                {
                    int item1 = i;
                    for (i = item1; i <= item1; i++)
                    {
                        for (int j = 1; j <= cells; j++)
                        {
                            var check = worksheet.Cells[i, j].Value?.ToString();

                            if (worksheet.Cells[i, j].Value == null)
                            {
                                if (j == 1)
                                {
                                    worksheet.Cells[i, j].Value = Convert.ToInt32(item.casenumber);
                                }
                                else if (j == 2)
                                {
                                    worksheet.Cells[i, j].Value = item.ClaimantFullName;
                                }
                                else if (j == 3)
                                {
                                    worksheet.Cells[i, j].Value = item.Service;
                                }
                                else if (j == 4)
                                {
                                    worksheet.Cells[i, j].Value = item.client;
                                }
                                else if (j == 5)
                                {
                                    worksheet.Cells[i, j].Value = item.company;
                                }
                                else if (j == 6)
                                {
                                    worksheet.Cells[i, j].Value = item.AppointmentDateStr;
                                }
                                else if (j == 7)
                                {
                                    worksheet.Cells[i, j].Value = item.Doctor;
                                }
                                else if (j == 8)
                                {
                                    worksheet.Cells[i, j].Value = item.Location;
                                }
                            }
                        }
                    }
                    item1 = i;
                }

                // save changes
                package.Save();
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in openXMLMailMerge/AppointmentLogExcelReport writing error Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                _return = false;
            }
            return _return;
        }


        public bool DailyStatusReportExcelReport(string TempPath, string FilePath, List<IMEDataReport2Updated> _IMEDataReport2, IMEReportFiltter IMEReportFiltter)
        {
            bool _return = true;
            try
            {
                FileInfo fileInfo = new FileInfo(TempPath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                // get number of rows in the sheet
                int rows = worksheet.Dimension.Rows; // 10
                int cells = worksheet.Dimension.Columns; // 10
                int i = 3;

                foreach (var item in _IMEDataReport2)
                {
                    int item1 = i;
                    for (i = item1; i <= item1; i++)
                    {
                        for (int j = 1; j <= cells; j++)
                        {
                            var check = worksheet.Cells[i, j].Value?.ToString();

                            if (worksheet.Cells[i, j].Value == null)
                            {
                                if (j == 1)
                                {
                                    worksheet.Cells[i, j].Value = item.AppointmentTimeString;
                                }
                                else if (j == 2)
                                {
                                    worksheet.Cells[i, j].Value = item.ClaimantFullName;
                                }
                                else if (j == 3)
                                {
                                    worksheet.Cells[i, j].Value = Convert.ToInt32(item.casenumber);
                                }
                                else if (j == 4)
                                {
                                    worksheet.Cells[i, j].Value = item.Service;
                                }
                                else if (j == 5)
                                {
                                    worksheet.Cells[i, j].Value = item.ShowsUp == false ? "" : "x";
                                }
                                else if (j == 6)
                                {
                                    worksheet.Cells[i, j].Value = item.NoShow == false ? "" : "x";
                                }
                                else if (j == 7)
                                {
                                    worksheet.Cells[i, j].Value = item.UnableToExamine == false ? "" : "x";
                                }
                                else if (j == 8)
                                {
                                    worksheet.Cells[i, j].Value = item.Comments;
                                }
                            }
                        }
                    }
                    item1 = i;
                }

                // save changes
                package.Save();
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in openXMLMailMerge/DailyStatusReportExcelReport writing error Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                _return = false;
            }
            return _return;
        }

        public bool ExportExcelReport(string TempPath, string FilePath, List<CaseReport> exports, IMEReportFiltter _IMEReportFiltter)
        {
            bool _return = true;
            try
            {
                FileInfo fileInfo = new FileInfo(TempPath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                // get number of rows in the sheet
                int rows = worksheet.Dimension.Rows; // 10
                int cells = worksheet.Dimension.Columns; // 10
                int i = 3;

                foreach (var item in exports)
                {
                    int item1 = i;
                    for (i = item1; i <= item1; i++)
                    {
                        for (int j = 1; j <= cells; j++)
                        {
                            var check = worksheet.Cells[i, j].Value?.ToString();

                            if (worksheet.Cells[i, j].Value == null)
                            {
                                if (j == 1)
                                {
                                    worksheet.Cells[i, j].Value = item.ClaimantName;
                                }
                                else if (j == 2)
                                {
                                    worksheet.Cells[i, j].Value = item.dateofexam;
                                }
                                else if (j == 3)
                                {
                                    worksheet.Cells[i, j].Value = item.Doctor;
                                }
                                else if (j == 4)
                                {
                                    worksheet.Cells[i, j].Value = item.DoctorLocation;
                                }
                                else if (j == 5)
                                {
                                    worksheet.Cells[i, j].Value = item.CaseChargeCount;
                                }
                            }
                        }
                    }
                    item1 = i;
                }

                // save changes
                package.Save();
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in openXMLMailMerge/ExportExcelReport writing error Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                _return = false;
            }
            return _return;
        }
    }
}