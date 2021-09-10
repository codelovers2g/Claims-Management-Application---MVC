using Buisness.Services;
using Common.Case.Models;
using Common.Case.Template;
using Common.QuickBooks;
using DataAccess.DataModels;
using DocumentsWebApp.CommonClass;
using DocumentWebApp.CommonClass;
using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Spire.Doc;
using Common.Master;
using System.Globalization;

namespace DocumentWebApp
{
    public class DocumentDownload
    {
        TemplateManager templateManager = new TemplateManager();
        DocumentManager DocumentManager = new DocumentManager();
        AddDocumentModal AddDocumentModals = new AddDocumentModal();


        public string demo()
        {
            return "From DocumentDownload";
        }

        public async Task<string> GenerateTemplates(AddDocumentModal AddDocumentModals)
        {
            int ProcessId = 0;
            var DPath = "";
            bool success = true;
            CaseManager caseManager = new CaseManager();
            string zipDownloadPath = "";
            try
            {
                bool status = false;
                int[] TemplateId = { 7, 12, 11 };
                string[] TemplateName = { "DISDL", "IME-1", "IME3" };

                string EmailFolderPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, "Uploaded\\EmailFolder\\");
                //Proces Status = initiated
                ProcessId = await caseManager.AddProcess(AddDocumentModals.CurrentUserId, AddDocumentModals.DateOffset, null);

                //Save CaseIds to Temporary Table
                await caseManager.AddProcessCasesToTemp(AddDocumentModals.BatchCaseId, ProcessId);

                bool isExists = System.IO.Directory.Exists(EmailFolderPath);

                var temp = Guid.NewGuid().ToString();

                if (!isExists)
                    System.IO.Directory.CreateDirectory(EmailFolderPath);

                string uploadedPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, "Uploaded\\EmailFolder\\" + temp + "\\");

                string NewFolderPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, "Uploaded\\EmailFolder\\" + temp + "\\");

                CustomLogger.LogError.GetLogger.Log(new Exception(), "5");

                bool NewFolderPathIsExists = System.IO.Directory.Exists(NewFolderPath);

                if (!NewFolderPathIsExists)
                    System.IO.Directory.CreateDirectory(NewFolderPath);

                CustomLogger.LogError.GetLogger.Log(new Exception(), "6");

                string MergedPdfDocument = "";
                List<string> MergedPdfDocuments = new List<string>();

                for (int i = 0; i < AddDocumentModals.BatchCaseId.Count; i++)
                {
                    AddDocumentModals.CaseId = AddDocumentModals.BatchCaseId[i];
                    var casedetail = caseManager.GetCaseById(AddDocumentModals.CaseId, AddDocumentModals);
                    string TargetDir = NewFolderPath + casedetail.FirstName + "_" + casedetail.LastName;
                    bool isExist = System.IO.Directory.Exists(TargetDir);
                    if (!isExist)
                        System.IO.Directory.CreateDirectory(TargetDir);
                    else
                    {
                        TargetDir = TargetDir + "  " + Guid.NewGuid().ToString();
                        System.IO.Directory.CreateDirectory(TargetDir);
                    }
                    //ProcessCase Status = Folder Created
                    status = await caseManager.AddCaseStatus(ProcessId, 6, AddDocumentModals.BatchCaseId[i]);
                    var CaseTypeLaibility = DocumentManager.CaseTypeRecordsById(AddDocumentModals.BatchCaseId[i]);
                    if (CaseTypeLaibility == true)
                    {
                        AddDocumentModals.TemplateId = 7;
                        AddDocumentModals.DocumentExt = ".DOCX";
                        AddDocumentModals.DocumentOrgName = "";
                        AddDocumentModals.DocumentName = "";
                        AddDocumentModals.TemplateName = "DISDL";
                        AddDocumentModals.TemplateTypeID = 2;
                        string DocPath = await AddChartPrepDocument(AddDocumentModals, "");
                        MergedPdfDocuments.Add(TargetDir + "\\" + AddDocumentModals.TemplateName + ".DOCX");

                        BlobStorage blobStorage = new BlobStorage();
                        string DocumentLocalPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, DocPath);
                        DocumentLocalPath = DocumentLocalPath.Replace("\\", "/");
                        using (var FileStream = System.IO.File.OpenRead(DocumentLocalPath))
                        {
                            string DocumentBlobPath = @"" + DocPath;
                            var tempUrl = await blobStorage.UploadFileOnBlob(FileStream, "schedulingapp", DocumentBlobPath);
                        }

                        DocPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, DocPath);
                        var Doc = DocPath.Replace("\\", "/");
                        string TargetPath = TargetDir + "\\" + AddDocumentModals.TemplateName + ".docx";
                        DPath = Doc + "   " + TargetPath;

                        await OpenXmlMailMerge.GetOpenXmlMailMerge.createfile(Doc, TargetPath);

                    }
                    else
                    {
                        for (int j = 0; j < TemplateId.Length; j++)
                        {
                            AddDocumentModals.TemplateId = TemplateId[j];
                            AddDocumentModals.DocumentExt = ".DOCX";
                            AddDocumentModals.DocumentOrgName = "";
                            AddDocumentModals.DocumentName = "";
                            AddDocumentModals.TemplateName = TemplateName[j];
                            AddDocumentModals.TemplateTypeID = 2;
                            string DocPath = await AddChartPrepDocument(AddDocumentModals, "");
                            MergedPdfDocuments.Add(TargetDir + "\\" + AddDocumentModals.TemplateName + ".DOCX");

                            BlobStorage blobStorage = new BlobStorage();
                            string DocumentLocalPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, DocPath);
                            DocumentLocalPath = DocumentLocalPath.Replace("\\", "/");
                            using (var FileStream = System.IO.File.OpenRead(DocumentLocalPath))
                            {
                                string DocumentBlobPath = @"" + DocPath;
                                var tempUrl = await blobStorage.UploadFileOnBlob(FileStream, "schedulingapp", DocumentBlobPath);
                            }

                            DocPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, DocPath);
                            var Doc = DocPath.Replace("\\", "/");
                            string TargetPath = TargetDir + "\\" + AddDocumentModals.TemplateName + ".docx";
                            DPath = Doc + "   " + TargetPath;

                            await OpenXmlMailMerge.GetOpenXmlMailMerge.createfile(Doc, TargetPath);

                        }
                    }
                    //ProcessCase Status = Files Created
                    status = await caseManager.AddCaseStatus(ProcessId, 7, AddDocumentModals.BatchCaseId[i]);

                    DocumentManager.UpdateChartPrep(AddDocumentModals, AddDocumentModals.BatchCaseId[i], AddDocumentModals.offset, CaseTypeLaibility);
                    //ProcessCase Status = Chart prepared
                    status = await caseManager.AddCaseStatus(ProcessId, 8, AddDocumentModals.BatchCaseId[i]);

                    //Delete CaseId from Temporary Table
                    await caseManager.DeleteProcessCasesFromTemp(AddDocumentModals.BatchCaseId[i], ProcessId);
                }
                //Proces Status = document generated
                status = await caseManager.UpdateProcessStatus(ProcessId, 2, "", "", temp, AddDocumentModals.DateOffset);

                string EmailFolder = NewFolderPath;
                string zipPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, "Uploaded\\EmailFolder\\" + temp + ".zip");

                zipDownloadPath = zipPath;
                ZipFile.CreateFromDirectory(EmailFolder, zipPath);

                //Proces Status = Zip Created
                status = await caseManager.UpdateProcessStatus(ProcessId, 3, "", "", temp, AddDocumentModals.DateOffset);

                // Upload Zip on Blob Storage
                BlobStorage zblobStorage = new BlobStorage();
                var zFileStream = System.IO.File.OpenRead(zipDownloadPath);
                string zDocumentBlobPath = "Uploaded/EmailFolder/" + temp + ".zip";
                var ziptempUrl = await zblobStorage.UploadFileOnBlob(zFileStream, "schedulingapp", zDocumentBlobPath);

                // Email Zip File
                //bool result = await ChartPrepSendEmail(zipDownloadPath, AddDocumentModals.CurrentUserEmail);

                // Documents Merged(Pdf)
                MergedPdfDocument = await GetMergeDocuments(MergedPdfDocuments, NewFolderPath, AddDocumentModals.BatchCaseId.Count(), 1);

                // Upload Zip on Blob Storage
                BlobStorage pdfblobStorage = new BlobStorage();
                var pdfFileStream = System.IO.File.OpenRead(MergedPdfDocument);
                string pdfDocumentBlobPath = "Uploaded/EmailFolder/" + temp + "/MergedDocument" + ".PDF";
                var pdftempUrl = await pdfblobStorage.UploadFileOnBlob(pdfFileStream, "schedulingapp", pdfDocumentBlobPath);

                // Documents Merged(Pdf)
                status = await caseManager.UpdateProcessStatus(ProcessId, 4, ziptempUrl, pdftempUrl, temp, AddDocumentModals.DateOffset);

                //Proces Status = Completed.
                status = await caseManager.UpdateProcessStatus(ProcessId, 5, ziptempUrl, pdftempUrl, temp, AddDocumentModals.DateOffset);

                return DPath;

            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in GenerateDocument writing error Function");
                //success = false;
                await caseManager.UpdateProcessFailure(ProcessId, ex.InnerException?.ToString(), 1);
                return DPath + " " + ex.ToString();
            }
        }

        public async Task RetryChart(AddDocumentModal AddDocumentModals)
        {
            int ProcessId = 0;
            CaseManager caseManager = new CaseManager();
            Process process = await caseManager.GetProcessById(AddDocumentModals.ProcessId);
            ProcessId = process.ProcessId;
            await caseManager.UpdateProcessFailure(ProcessId, null, 2);
            AddDocumentModals.BatchCaseId = await caseManager.GetBactchIdsByProcessId(ProcessId);
            string zipDownloadPath = "";

            try
            {
                bool status = false;
                int[] TemplateId = { 7, 12, 11 };
                string[] TemplateName = { "DISDL", "IME-1", "IME3" };

                string EmailFolderPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, "Uploaded\\EmailFolder\\");

                //Save CaseIds to Temporary Table
                await caseManager.AddProcessCasesToTemp(AddDocumentModals.BatchCaseId, ProcessId);

                bool isExists = System.IO.Directory.Exists(EmailFolderPath);

                var temp = "";
                if (process.WorkingFolder_ExportTotal != null)
                {
                    temp = process.WorkingFolder_ExportTotal;
                }
                else
                {
                    temp = Guid.NewGuid().ToString();
                }

                if (!isExists)
                    System.IO.Directory.CreateDirectory(EmailFolderPath);

                string GuidFolderPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, "Uploaded\\EmailFolder\\" + temp + "\\");

                bool NewFolderPathIsExists = System.IO.Directory.Exists(GuidFolderPath);

                if (!NewFolderPathIsExists)
                    System.IO.Directory.CreateDirectory(GuidFolderPath);
                string MergedPdfDocument = "";
                List<string> MergedPdfDocuments = new List<string>();
                if (AddDocumentModals.BatchCaseId != null)
                {
                    for (int i = 0; i < AddDocumentModals.BatchCaseId.Count; i++)
                    {
                        AddDocumentModals.CaseId = AddDocumentModals.BatchCaseId[i];
                        var casedetail = caseManager.GetCaseById(AddDocumentModals.CaseId, AddDocumentModals);
                        string TargetDir = GuidFolderPath + casedetail.FirstName + "_" + casedetail.LastName;
                        bool isExist = System.IO.Directory.Exists(TargetDir);
                        if (!isExist)
                            System.IO.Directory.CreateDirectory(TargetDir);
                        else
                        {
                            TargetDir = TargetDir + "  " + Guid.NewGuid().ToString();
                            System.IO.Directory.CreateDirectory(TargetDir);
                        }
                        //ProcessCase Status = Folder Created
                        status = await caseManager.AddCaseStatus(ProcessId, 6, AddDocumentModals.BatchCaseId[i]);
                        var CaseTypeLaibility = DocumentManager.CaseTypeRecordsById(AddDocumentModals.BatchCaseId[i]);
                        if (CaseTypeLaibility == true)
                        {
                            AddDocumentModals.TemplateId = 7;
                            AddDocumentModals.DocumentExt = ".DOCX";
                            AddDocumentModals.DocumentOrgName = "";
                            AddDocumentModals.DocumentName = "";
                            AddDocumentModals.TemplateName = "DISDL";
                            AddDocumentModals.TemplateTypeID = 2;
                            string DocPath = await AddChartPrepDocument(AddDocumentModals, "");
                            MergedPdfDocuments.Add(TargetDir + "\\" + AddDocumentModals.TemplateName + ".DOCX");

                            BlobStorage blobStorage = new BlobStorage();
                            string DocumentLocalPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, DocPath);
                            DocumentLocalPath = DocumentLocalPath.Replace("\\", "/");
                            using (var FileStream = System.IO.File.OpenRead(DocumentLocalPath))
                            {
                                string DocumentBlobPath = @"" + DocPath;
                                var tempUrl = await blobStorage.UploadFileOnBlob(FileStream, "schedulingapp", DocumentBlobPath);
                            }

                            DocPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, DocPath);
                            var Doc = DocPath.Replace("\\", "/");
                            string TargetPath = TargetDir + "\\" + AddDocumentModals.TemplateName + ".docx";

                            await OpenXmlMailMerge.GetOpenXmlMailMerge.createfile(Doc, TargetPath);

                        }
                        else
                        {
                            for (int j = 0; j < TemplateId.Length; j++)
                            {
                                AddDocumentModals.TemplateId = TemplateId[j];
                                AddDocumentModals.DocumentExt = ".DOCX";
                                AddDocumentModals.DocumentOrgName = "";
                                AddDocumentModals.DocumentName = "";
                                AddDocumentModals.TemplateName = TemplateName[j];
                                AddDocumentModals.TemplateTypeID = 2;
                                string DocPath = await AddChartPrepDocument(AddDocumentModals, "");
                                MergedPdfDocuments.Add(TargetDir + "\\" + AddDocumentModals.TemplateName + ".DOCX");

                                BlobStorage blobStorage = new BlobStorage();
                                string DocumentLocalPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, DocPath);
                                DocumentLocalPath = DocumentLocalPath.Replace("\\", "/");
                                using (var FileStream = System.IO.File.OpenRead(DocumentLocalPath))
                                {
                                    string DocumentBlobPath = @"" + DocPath;
                                    var tempUrl = await blobStorage.UploadFileOnBlob(FileStream, "schedulingapp", DocumentBlobPath);
                                }
                                DocPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, DocPath);
                                var Doc = DocPath.Replace("\\", "/");
                                string TargetPath = TargetDir + "\\" + AddDocumentModals.TemplateName + ".docx";

                                await OpenXmlMailMerge.GetOpenXmlMailMerge.createfile(Doc, TargetPath);

                            }
                        }
                        //ProcessCase Status = Files Created
                        status = await caseManager.AddCaseStatus(ProcessId, 7, AddDocumentModals.BatchCaseId[i]);
                        DocumentManager.UpdateChartPrep(AddDocumentModals, AddDocumentModals.BatchCaseId[i], AddDocumentModals.offset, CaseTypeLaibility);

                        //ProcessCase Status = Chart prepared
                        status = await caseManager.AddCaseStatus(ProcessId, 8, AddDocumentModals.BatchCaseId[i]);

                        //Delete CaseId from Temporary Table
                        await caseManager.DeleteProcessCasesFromTemp(AddDocumentModals.BatchCaseId[i], ProcessId);
                    }
                }
                //Proces Status = document generated
                status = await caseManager.UpdateProcessStatus(ProcessId, 2, "", "", temp, AddDocumentModals.DateOffset);

                string EmailFolder = GuidFolderPath;
                string zipPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, "Uploaded\\EmailFolder\\" + temp + ".zip");

                zipDownloadPath = zipPath;

                if (File.Exists(zipDownloadPath) != true)
                    ZipFile.CreateFromDirectory(EmailFolder, zipPath);

                //Proces Status = Zip Created
                status = await caseManager.UpdateProcessStatus(ProcessId, 3, "", "", temp, AddDocumentModals.DateOffset);

                // Upload Zip on Blob Storage
                BlobStorage zblobStorage = new BlobStorage();
                var zFileStream = System.IO.File.OpenRead(zipDownloadPath);
                string zDocumentBlobPath = "Uploaded/EmailFolder/" + temp + ".zip";
                var ziptempUrl = await zblobStorage.UploadFileOnBlob(zFileStream, "schedulingapp", zDocumentBlobPath);

                // Email Zip File
                //bool result = await ChartPrepSendEmail(zipDownloadPath, AddDocumentModals.CurrentUserEmail);

                // Documents Merged(Pdf)
                MergedPdfDocument = await GetMergeDocuments(MergedPdfDocuments, GuidFolderPath, AddDocumentModals.BatchCaseId.Count(), 1);

                // Upload Zip on Blob Storage
                BlobStorage pdfblobStorage = new BlobStorage();
                var pdfFileStream = System.IO.File.OpenRead(MergedPdfDocument);
                string pdfDocumentBlobPath = "Uploaded/EmailFolder/" + temp + ".zip";
                var pdftempUrl = await pdfblobStorage.UploadFileOnBlob(zFileStream, "schedulingapp", zDocumentBlobPath);

                // Documents Merged(Pdf)
                status = await caseManager.UpdateProcessStatus(ProcessId, 4, ziptempUrl, pdftempUrl, temp, AddDocumentModals.DateOffset);

                //Proces Status = Completed.
                status = await caseManager.UpdateProcessStatus(ProcessId, 5, ziptempUrl, pdftempUrl, temp, AddDocumentModals.DateOffset);



            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in GenerateDocument writing error Function");
                await caseManager.UpdateProcessFailure(ProcessId, ex.InnerException?.ToString(), 1);
            }
        }

        public async Task GenerateChecks(AddDocumentModal AddDocumentModals)
        {
            CaseManager caseManager = new CaseManager();
            List<Vendors> vendorDetail = null;
            int ProcessId = AddDocumentModals.ProcessId;
            int TotalPayees = 0;
            try
            {
                vendorDetail = caseManager.GetPayeesChecksDetails(AddDocumentModals.CaseIdShow);

                bool status = false;
                string checksFolderPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, "Uploaded\\Checks\\");
                var processFolderName = Guid.NewGuid().ToString();

                bool isExists = System.IO.Directory.Exists(checksFolderPath);
                if (!isExists)
                    System.IO.Directory.CreateDirectory(checksFolderPath);

                string processPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, "Uploaded\\Checks\\" + processFolderName + "\\");

                bool processPathExists = System.IO.Directory.Exists(processPath);
                if (!processPathExists)
                    System.IO.Directory.CreateDirectory(processPath);

                //Proces Status = document generated
                status = await caseManager.UpdateProcessStatus(ProcessId, 15, "", "", TotalPayees.ToString(), AddDocumentModals.DateOffset);

                string MergedPdfDocument = "";
                List<string> MergedPdfDocuments = new List<string>();

                foreach (var item in vendorDetail)
                {
                    if (item.PayeeCharges != 0)
                    {

                        AddDocumentModals.TemplateId = 42;
                        AddDocumentModals.DocumentExt = ".DOCX";
                        AddDocumentModals.DocumentOrgName = "";
                        AddDocumentModals.DocumentName = "";
                        AddDocumentModals.TemplateName = "Checks";
                        AddDocumentModals.TemplateTypeID = 2;
                        AddDocumentModals.ProcessFolder = processFolderName;

                        // NYCTemplate 
                        NYCTemplate _NYCTemplate = new NYCTemplate();
                        _NYCTemplate.DateNow = DateTime.UtcNow.ToString("MM/dd/yyyy");
                        _NYCTemplate.ServiceProviderFullName = item.DisplayName;
                        _NYCTemplate.ServiceProviderCredential = item.PayeeCredential;
                        _NYCTemplate.ServiceProviderMailingAddress = item.MailingAddress;
                        _NYCTemplate.Vendor_Charges = Convert.ToDecimal(item.PayeeCharges).ToString("#,##0.00");
                        _NYCTemplate.Vendor_Charges_Word = item.PayeeChargesWords;

                        string DocPath = await AddCheckDocument(AddDocumentModals, _NYCTemplate);

                        if (DocPath != "Error")
                            TotalPayees++;

                        MergedPdfDocuments.Add(AddDocumentModals.LocalServerPath + DocPath);

                        BlobStorage blobStorage = new BlobStorage();
                        string DocumentLocalPath = AddDocumentModals.LocalServerPath + DocPath.Replace("\\", "/");
                        var tempUrl = "";
                        using (var FileStream = System.IO.File.OpenRead(DocumentLocalPath))
                        {
                            string DocumentBlobPath = @"" + DocPath;
                            tempUrl = await blobStorage.UploadFileOnBlob(FileStream, "schedulingapp", DocumentBlobPath);
                            FileStream.Close();
                        }

                        DocPath = System.IO.Path.Combine(AddDocumentModals.LocalServerPath, DocPath);
                        var Doc = DocPath.Replace("\\", "/");


                        //Update CASES History and Status for Case

                        List<CaseDetailsModel> caseDetails = caseManager.GetFilteredExport(item.CaseIdList);

                        foreach (var caseDetail in caseDetails)
                        {
                            QuickBooksManager QuickBooksManager = new QuickBooksManager();
                            UserModal _UserModal = new UserModal();
                            _UserModal.UserId = AddDocumentModals.CurrentUserId;
                            var doctorOrPayeeApiId = QuickBooksManager.GetDoctorOrpayeeApiId(caseDetail.doctorIds, caseDetail.doctorLocationIds);
                            bool res = QuickBooksManager.SaveCaseChargesApiId(caseDetail.caseIds, "", caseDetail.caseVendorCharges);
                            if (caseDetail.caseVendorCharges != 0)
                                QuickBooksManager.SaveCaseChargesHistory(caseDetail.caseIds, caseDetail.caseNumbers, caseDetail.caseVendorCharges, item.DisplayName, _UserModal);
                            await caseManager.DeleteProcessCasesFromTemp(caseDetail.caseIds, ProcessId);
                        }

                        //Add Checks Info

                        await caseManager.AddChecks(item, tempUrl, AddDocumentModals.checkIdSequence);
                        AddDocumentModals.checkIdSequence++;
                    }
                    else
                    {
                        TotalPayees++;
                        List<CaseDetailsModel> caseDetails = caseManager.GetFilteredExport(item.CaseIdList);

                        foreach (var caseDetail in caseDetails)
                        {
                            QuickBooksManager QuickBooksManager = new QuickBooksManager();
                            UserModal _UserModal = new UserModal();
                            _UserModal.UserId = AddDocumentModals.CurrentUserId;
                            var doctorOrPayeeApiId = QuickBooksManager.GetDoctorOrpayeeApiId(caseDetail.doctorIds, caseDetail.doctorLocationIds);
                            bool res = QuickBooksManager.SaveCaseChargesApiId(caseDetail.caseIds, "", caseDetail.caseVendorCharges);
                            await caseManager.DeleteProcessCasesFromTemp(caseDetail.caseIds, ProcessId);
                        }
                    }
                }

                //Proces Status = document generated
                status = await caseManager.UpdateProcessStatus(ProcessId, 16, "", "", TotalPayees.ToString(), AddDocumentModals.DateOffset);

                if(MergedPdfDocuments.Count != 0)
                {
                    // Documents Merged(Pdf)
                    MergedPdfDocument = await GetMergeDocuments(MergedPdfDocuments, checksFolderPath + processFolderName, AddDocumentModals.CaseIdShow.Count(), 2);

                    // Upload Zip on Blob Storage
                    BlobStorage pdfblobStorage = new BlobStorage();
                    var pdfFileStream = System.IO.File.OpenRead(MergedPdfDocument);
                    string pdfDocumentBlobPath = "Uploaded/Checks/" + processFolderName + "/ChecksCombined" + ".PDF";
                    var pdftempUrl = await pdfblobStorage.UploadFileOnBlob(pdfFileStream, "schedulingapp", pdfDocumentBlobPath);
                    pdfFileStream.Close();
                    //Documents Merged(Pdf)
                    status = await caseManager.UpdateProcessStatus(ProcessId, 17, "", pdftempUrl, TotalPayees.ToString(), AddDocumentModals.DateOffset);
                }
                else
                {
                    //Export Completed
                    status = await caseManager.UpdateProcessStatus(ProcessId, 17, "", "", TotalPayees.ToString(), AddDocumentModals.DateOffset);
                }
                

            }
            catch (Exception ex)
            {
                await caseManager.DeleteProcessFromTemp(ProcessId);
                await caseManager.UpdateProcessFailure(ProcessId, ex.InnerException?.ToString(), 1);
                CustomLogger.LogError.GetLogger.Log(ex, "Error in GenerateCheckDocument writing error Function");
                //success = false;
            }
        }


        public async Task<bool> ChartPrepSendEmail(string zipDownloadPath, string email)
        {
            SendDocMail Mail = new SendDocMail();

            //Mail.To to = new List<string>();
            List<string> toList = new List<string>();
            //toList.Add(email);

            List<string> attachment = new List<string>();
            attachment.Add(zipDownloadPath);

            Mail.To = toList;
            Mail.Attchment = attachment;
            Mail.Subject = "Zip File for Chart Prep Process By ";
            Mail.Body = "PFA, \n\n\nThanks and Regards\nUtopia Claims";

            bool result = await SendMailWithDocmentsZip(Mail);

            return result;
        }

        public async Task<string> AddChartPrepDocument(AddDocumentModal AddDocumentModal, string docPath)
        {
            string _return = null;
            try
            {
                DocumentManager DocumentManager = new DocumentManager();
                TemplateManager templateManager = new TemplateManager();
                bool isSavedSuccessfully = true;

                if (AddDocumentModal.TemplateId != null && AddDocumentModal.TemplateId != 0)
                {
                    if (AddDocumentModal.DocumentOrgName == null || AddDocumentModal.DocumentOrgName == "" || AddDocumentModal.DocumentOrgName == "undefined")
                    {
                        AddDocumentModal.DocumentOrgName = AddDocumentModal.TemplateName + AddDocumentModal.DocumentExt;
                        AddDocumentModal.DocumentName = (AddDocumentModal.TemplateName ?? "") + "_" + Guid.NewGuid().ToString() + AddDocumentModal.DocumentExt;
                    }
                    //var TempaltePath = new DirectoryInfo(string.Format("{0}Documents\\Templates\\" + AddDocumentModal.TemplateName + AddDocumentModal.DocumentExt, HttpContext.Current.Server.MapPath(@"\"))).ToString();
                    var TempaltePath = new DirectoryInfo(string.Format("{0}Documents\\Templates\\" + AddDocumentModal.TemplateName + AddDocumentModal.DocumentExt, AddDocumentModal.LocalServerPath)).ToString();
                    if (System.IO.File.Exists(TempaltePath.ToString()))
                    {
                        var originalDirectory = new DirectoryInfo(string.Format("{0}Uploaded\\Documents\\", AddDocumentModal.LocalServerPath));
                        string pathString = System.IO.Path.Combine(originalDirectory.ToString(), AddDocumentModal.CaseId + "\\Document");
                        bool isExists = System.IO.Directory.Exists(pathString);
                        if (!isExists)
                            System.IO.Directory.CreateDirectory(pathString);
                        var FilePath = new DirectoryInfo(string.Format("{0}Uploaded\\Documents\\" + AddDocumentModal.CaseId + "\\Document\\" + AddDocumentModal.DocumentName, AddDocumentModal.LocalServerPath)).ToString();
                        var FilePath2 = new DirectoryInfo(string.Format("{0}Uploaded\\Documents\\" + AddDocumentModal.CaseId + "\\Document\\" + "New" + AddDocumentModal.DocumentName, AddDocumentModal.LocalServerPath)).ToString();
                        if (!System.IO.File.Exists(FilePath))
                        {
                            if (!System.IO.Directory.Exists(string.Format("{0}Documents\\Temp\\", AddDocumentModal.LocalServerPath)))
                            {
                                System.IO.Directory.CreateDirectory(string.Format("{0}Documents\\Temp\\", AddDocumentModal.LocalServerPath));
                            }
                            var TempPath = new DirectoryInfo(string.Format("{0}Documents\\Temp\\" + AddDocumentModal.DocumentName, AddDocumentModal.LocalServerPath)).ToString();
                            System.IO.File.Copy(TempaltePath, TempPath);
                            System.IO.File.Copy(TempaltePath, FilePath2);
                            Dictionary<string, string> pDictionaryMerge = new Dictionary<string, string>();
                            Dictionary<string, string> UniversalField = null;
                            TemplateManager _TemplateManager = new TemplateManager();
                            if (AddDocumentModal.TemplateName == TemplateName.IME3.ToString())
                            {
                                pDictionaryMerge = await _TemplateManager.GetIME3Template(AddDocumentModal.CaseId, Convert.ToInt32(AddDocumentModal.ServiceId));
                            }
                            else if (AddDocumentModal.TemplateName == TemplateName.DISDL.ToString())
                            {
                                pDictionaryMerge = await _TemplateManager.GetDISDLTemplate(AddDocumentModal.CaseId, Convert.ToInt32(AddDocumentModal.ServiceId));
                            }
                            else if (AddDocumentModal.TemplateName == "IME-1")
                            {
                                pDictionaryMerge = await _TemplateManager.GetIME1Template(AddDocumentModal.CaseId, Convert.ToInt32(AddDocumentModal.ServiceId));
                            }
                            else
                            {
                                pDictionaryMerge = await _TemplateManager.GetUniversalTemplate(AddDocumentModal.CaseId, Convert.ToInt32(AddDocumentModal.ServiceId));
                                UniversalField = pDictionaryMerge;
                            }
                            //Copying Universal field to replace template text field which is used from template form copy paste option - discussed with sam
                            if (UniversalField == null)
                            {
                                UniversalField = await _TemplateManager.GetUniversalTemplate(AddDocumentModal.CaseId, Convert.ToInt32(AddDocumentModal.ServiceId));
                            }
                            /// Gets all the fields inside HealthCare Practioner group
                            var ArrMergeFieldsInTableStart = GetFieldsInHealthCare(FilePath2);

                            NYCTemplate nYCTemplateObj = templateManager.GetCaseDetails(AddDocumentModal.CaseId, Convert.ToInt32(AddDocumentModal.ServiceId), ArrMergeFieldsInTableStart);

                            bool IsMergeSuccess = false;
                            if (AddDocumentModal.TemplateName == "IME-1")
                            {
                                CaseManager caseManager = new CaseManager();
                                IME4HeaderData iMEHearderData = await caseManager.GetIME4HeaderData(AddDocumentModal.CaseId);
                                string TemplatePath = System.IO.Path.Combine(AddDocumentModal.LocalServerPath, "Uploaded\\Documents\\" + AddDocumentModal.CaseId + "\\Document\\") + AddDocumentModal.DocumentName;
                                var Temp = TempPath.Replace("\\", "/");
                                await OpenXmlMailMerge.GetOpenXmlMailMerge.copyfile(Temp, FilePath);
                                string HeaderPath = System.IO.Path.Combine(AddDocumentModal.LocalServerPath, "Uploaded\\IME1-Header.docx");
                                await OpenXmlWordHelpers.ChangeHeader(HeaderPath, iMEHearderData);
                                await OpenXmlWordHelpers.PrependHeader(HeaderPath, FilePath);
                                IsMergeSuccess = await OpenXmlMailMerge.GetOpenXmlMailMerge.OpenXmlMerge(TempPath, FilePath, pDictionaryMerge, nYCTemplateObj, UniversalField);
                            }
                            else
                            {
                                await OpenXmlMailMerge.GetOpenXmlMailMerge.copyfile(TempPath, FilePath);
                                IsMergeSuccess = await OpenXmlMailMerge.GetOpenXmlMailMerge.OpenXmlMerge(TempPath, FilePath, pDictionaryMerge, nYCTemplateObj, UniversalField);
                            }

                            if (IsMergeSuccess == false)
                            {
                                return "Error";
                            }
                        }
                    }
                }

                var tempDocument = await DocumentManager.AddDocument(AddDocumentModal);

                if (tempDocument.Success)
                {
                    _return = "Uploaded/Documents/" + tempDocument.CaseId.ToString() + "/Document/" + tempDocument.DocumentName;
                }

            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in AddDocument Function");
                try
                {
                    if (ex != null)
                    {
                        CustomLogger.LogError.GetLogger.Log(ex, "Error in AddDocument Function.");
                        QuickBookLogCQB _QuickBookLogCQB = new QuickBookLogCQB();
                        _QuickBookLogCQB.AccessToken = " "; //acces_token variable is defined in AppController as this controller is inherited from AppController
                        _QuickBookLogCQB.SyncRequest = "123124124";
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
                    CustomLogger.LogError.GetLogger.Log(ex1, "Error in AddDocument writing error Function");
                }
                _return = "Error";
            }
            return _return;
        }
        public async Task<string> AddCheckDocument(AddDocumentModal AddDocumentModal, NYCTemplate _NYCTemplate)
        {
            string _return = null;
            try
            {
                DocumentManager DocumentManager = new DocumentManager();
                TemplateManager templateManager = new TemplateManager();
                bool isSavedSuccessfully = true;

                if (AddDocumentModal.TemplateId != null && AddDocumentModal.TemplateId != 0)
                {
                    if (AddDocumentModal.DocumentOrgName == null || AddDocumentModal.DocumentOrgName == "" || AddDocumentModal.DocumentOrgName == "undefined")
                    {
                        AddDocumentModal.DocumentOrgName = AddDocumentModal.TemplateName + AddDocumentModal.DocumentExt;
                        AddDocumentModal.DocumentName = (AddDocumentModal.TemplateName ?? "") + "_" + Guid.NewGuid().ToString() + AddDocumentModal.DocumentExt;
                    }
                    var TempaltePath = new DirectoryInfo(string.Format("{0}Documents\\Templates\\" + AddDocumentModal.TemplateName + AddDocumentModal.DocumentExt, AddDocumentModal.LocalServerPath)).ToString();
                    if (System.IO.File.Exists(TempaltePath.ToString()))
                    {
                        var originalDirectory = new DirectoryInfo(string.Format("{0}Uploaded\\Checks\\", AddDocumentModal.ProcessFolder));
                        var FilePath = new DirectoryInfo(string.Format("{0}Uploaded\\Checks\\" + AddDocumentModal.ProcessFolder + "\\" + _NYCTemplate.ServiceProviderFullName + ".docx", AddDocumentModal.LocalServerPath));
                        var FilePath2 = new DirectoryInfo(string.Format("{0}Uploaded\\Checks\\" + AddDocumentModal.ProcessFolder + "\\" + "New" + _NYCTemplate.ServiceProviderFullName + ".docx", AddDocumentModal.LocalServerPath));
                        if (!System.IO.File.Exists(FilePath.ToString()))
                        {
                            if (!System.IO.Directory.Exists(string.Format("{0}Documents\\Temp\\", AddDocumentModal.LocalServerPath)))
                            {
                                System.IO.Directory.CreateDirectory(string.Format("{0}Documents\\Temp\\", AddDocumentModal.LocalServerPath));
                            }
                            var TempPath = new DirectoryInfo(string.Format("{0}Documents\\Temp\\" + AddDocumentModal.DocumentName, AddDocumentModal.LocalServerPath)).ToString();
                            System.IO.File.Copy(TempaltePath, TempPath);
                            System.IO.File.Copy(TempaltePath, FilePath2.ToString());
                            Dictionary<string, string> pDictionaryMerge = new Dictionary<string, string>();
                            Dictionary<string, string> UniversalField = null;
                            TemplateManager _TemplateManager = new TemplateManager();
                            pDictionaryMerge = await _TemplateManager.GetChecksFields(_NYCTemplate);
                            UniversalField = pDictionaryMerge;
                            //Copying Universal field to replace template text field which is used from template form copy paste option - discussed with sam
                            if (UniversalField == null)
                            {
                                UniversalField = await _TemplateManager.GetChecksFields(_NYCTemplate);
                            }
                            bool IsMergeSuccess = false;
                            await OpenXmlMailMerge.GetOpenXmlMailMerge.copyfile(TempPath, FilePath.ToString());
                            IsMergeSuccess = await OpenXmlMailMerge.GetOpenXmlMailMerge.OpenXmlMerge(TempPath, FilePath.ToString(), pDictionaryMerge, _NYCTemplate, UniversalField);

                            if (IsMergeSuccess == false)
                            {
                                return "Error";
                            }
                        }
                    }
                }
                _return = "Uploaded/Checks/" + AddDocumentModal.ProcessFolder + "/" + _NYCTemplate.ServiceProviderFullName + ".docx";
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in AddDocument Function");
                _return = "Error";
            }
            return _return;
        }
        private string[] GetFieldsInHealthCare(string TemplatePath)
        {
            var _return = new string[] { "" };
            try
            {
                Document document = new Document();
                document.LoadFromFile(TemplatePath);
                var groups = document.MailMerge.GetMergeGroupNames();
                if (groups.Length > 0)
                {
                    string[] Names = document.MailMerge.GetMergeFieldNames(groups[0]);
                    _return = Names;
                }
                else
                {
                    _return = null;
                }
            }
            catch (Exception ex)
            {
                //throw;
            }
            return _return;
        }
        public async Task<bool> SendMailWithDocmentsZip(SendDocMail SendDocMail)
        {
            bool _return = true;
            try
            {
                CaseManager caseManager = new CaseManager();
                //string UserEmail = caseManager.GetCurrentUser();
                SendDocMail.To[0] = "anjali.step2gen@gmail.com";
                //SendDocMail.To[0] = UserEmail;
                SendDocMail.Body = SendDocMail.Body.Replace("\n", "<br>");
                await EmailConfiguration.GetEmailConfiguration.SendMailWithDocmentsZip(string.Join(",", SendDocMail.To.Distinct().ToList()), SendDocMail.Subject, SendDocMail.Body, string.Join(",", SendDocMail.Attchment.Distinct().ToList()));
            }
            catch (Exception ex)
            {

                CustomLogger.LogError.GetLogger.Log(ex, "Error in AddDocument Function");
                _return = false;
            }
            return _return;
        }
        public async Task<string> GetMergeDocuments(List<string> DocumentsPaths, string ProcessFolder, int TotalCases, int offset)
        {
            string MergeDocumentPath = string.Empty;
            List<string> GeneratedDocumentPath = new List<string>();
            List<DocumentListing> DocumentListing = new List<DocumentListing>();

            try
            {
                foreach (var Document in DocumentsPaths)
                {
                    string MergedDocumentPath = System.IO.Path.Combine(ProcessFolder, "Pdf");
                    string tempDocument = "";
                    try
                    {
                        bool isExists = System.IO.Directory.Exists(MergedDocumentPath);
                        if (!isExists)
                            System.IO.Directory.CreateDirectory(MergedDocumentPath);
                        string PdfPath = MergedDocumentPath + System.IO.Path.GetFileNameWithoutExtension(Document) + "_" + Guid.NewGuid().ToString() + ".pdf";

                        if (!System.IO.File.Exists(PdfPath))
                        {
                            string CopyPath = Path.Combine(MergedDocumentPath, Guid.NewGuid().ToString() + System.IO.Path.GetExtension(Document));
                            if (System.IO.File.Exists(Document))
                            {
                                System.IO.File.Copy(Document, CopyPath);
                                if (CovertDocToPDF(CopyPath, PdfPath))
                                {
                                    tempDocument = PdfPath;
                                }
                            }
                        }
                        else
                        {
                            tempDocument = PdfPath;
                        }
                        GeneratedDocumentPath.Add(tempDocument);
                    }
                    catch (Exception ex)
                    {
                        CustomLogger.LogError.GetLogger.Log(ex, "Error in CovertDocToPDF Function,");
                    }
                }
                MergeDocumentPath = await MergeDocuments(GeneratedDocumentPath, ProcessFolder, TotalCases, offset);
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in GetDocumentList Function");
            }
            return MergeDocumentPath;
        }

        public async Task<string> MergeDocuments(List<string> GeneratedDocumentPath, string ProcessFolder, int TotalCases, int offset)
        {
            string[] files = GeneratedDocumentPath.ToArray();

            string mergeFilePath = "";
            if (offset == 1)
                mergeFilePath = Path.Combine(ProcessFolder.ToString(), "ChartPrep_" + TotalCases + "Cases.pdf");
            else
                mergeFilePath = Path.Combine(ProcessFolder.ToString(), "Checks_" + TotalCases + "Vendors.pdf");

            PdfDocumentBase doc = PdfDocument.MergeFiles(files);
            doc.Save(mergeFilePath);
            return mergeFilePath;
        }
        public bool CovertDocToPDF(String Docfile, string PDFFile)
        {
            bool _return = false;
            try
            {
                //Force clean up
                GC.Collect();
                if (File.Exists(Docfile))
                {
                    System.Threading.Thread.Sleep(1500);

                    Document doc = new Document();
                    doc.LoadFromFile(Docfile);
                    ToPdfParameterList tpl = new ToPdfParameterList
                    {
                        UsePSCoversion = true//azure
                    };
                    doc.SaveToFile(PDFFile, tpl);
                    /// code bellow not working on azure
                    //Spire.Doc.Document document = new Spire.Doc.Document(Docfile, FileFormat.Auto);
                    //document.SaveToFile(PDFFile, FileFormat.PDF);
                    _return = true;
                }
                else
                {
                    Exception exCustom = new Exception("File Doesn't exist : " + Docfile);
                    throw exCustom;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return _return;
        }
    }
}
