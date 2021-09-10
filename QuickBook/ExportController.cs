using Buisness.Services;
using Common.Case.Models;
using Common.Master;
using Common.QuickBooks;
using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.DataService;
using Intuit.Ipp.QueryFilter;
using Newtonsoft.Json;
using SchedulingApp.CommonClass;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace SchedulingApp.Areas.Accounts.Controllers
{
    public class ExportController : AppController
    {
        // GET: Accounts/Export
        CommonManager CommonManager;
        QuickBooksManager QuickBooksManager;
        ReportManager ReportManager;
        CaseManager CaseManager;
        public ExportController()
        {
            CommonManager = new CommonManager();
            QuickBooksManager = new QuickBooksManager();
            ReportManager = new ReportManager();
            CaseManager = new CaseManager();
        }


        //public JsonResult ExportToQuickBooks(CaseDetailsModel caseDetails)
        public async Task<JsonResult> ExportToQuickBooks(List<CaseDetailsModel> caseDetails)
        {
            CaseManager caseManager = new CaseManager();
            string realmId = Session["realmId"].ToString();
            if (realmId == "")
            {
                RefreshToken();
            }
            UserModal _UserModalobj = (UserModal)Session[SessionEnum.loggedinuser.ToString()];
            List<CaseDetailsModel> caseDetail = null;
            List<int> caseIds = new List<int>();
            int ProcessId = 0;
            try
            {
                caseDetail = CaseManager.GetFilteredExport(caseDetails);
                if (_UserModalobj != null)
                {
                    if (caseDetails != null)
                        foreach (var item in caseDetail)
                            caseIds.Add(item.caseIds);
                    ProcessId = await CaseManager.AddProcess(_UserModalobj.UserId, 0, caseDetails.Count.ToString());
                    await CaseManager.AddProcessCasesToTemp(caseIds, ProcessId);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            var principal = User as ClaimsPrincipal;
            System.Threading.Tasks.Task.Run(() =>
            {
                RedirectExportToQuickBooks(caseDetail, realmId, _UserModalobj, principal, ProcessId);
            }).ConfigureAwait(false);
            return Json("", JsonRequestBehavior.AllowGet);
        }

        public async System.Threading.Tasks.Task RedirectExportToQuickBooks(List<CaseDetailsModel> caseDetails, string realmId, UserModal _UserModalobj, ClaimsPrincipal principal, int ProcessId)
        {
            CaseManager CaseManager = new CaseManager();
            decimal result = 0;
            //int ProcessId = 0;
            //Make QBO api calls using .Net SDK
            try
            {
                int accountPayablesSynced = 0, accountBillsSynced = 0, accountBillPaymentsSynced = 0, deletedBillPaymentsSynced;
                var d = Request.Params["SyncDate"];
                var uptoDate = Convert.ToDateTime(d);
                bool status = true;
                string getUptoDate = Convert.ToString(Request.Params["SyncDate"]);
                //DateTime date = DateTime.Parse(getUptoDate);
                var uptoDate1 = Convert.ToDateTime(getUptoDate);
                string accessToken = string.Empty;
                if (realmId != null)
                {
                    if (_UserModalobj != null)
                    {
                        //Proces Status = Initiated
                        //ProcessId = await CaseManager.AddProcess(_UserModalobj.UserId, 0, caseDetails.Count.ToString());

                        //Initialize OAuth2RequestValidator and ServiceContext
                        ServiceContext serviceContext = base.IntializeContext(realmId, principal);

                        //For checking payee and doctor API_Id(null) on quickbooks
                        foreach (var caseDetail in caseDetails)
                        {
                            var doctorOrPayeeApiId = QuickBooksManager.GetDoctorOrpayeeApiId(caseDetail.doctorIds, caseDetail.doctorLocationIds);

                            QueryService<Vendor> vendorQuerySvc = new QueryService<Vendor>(serviceContext);
                            var vendorInfo = vendorQuerySvc.ExecuteIdsQuery("Select * From vendor where Id = '" + doctorOrPayeeApiId + "'").FirstOrDefault();
                            if (vendorInfo == null)
                            {
                                await CaseManager.UpdateDoctorsNotAvailableInQBs(doctorOrPayeeApiId);
                            }

                        }

                        //Add doctor and payeee on QuickBook
                        await AdddoctorsandPayeestoQBs(serviceContext, ProcessId, _UserModalobj);

                        #region Get/Sync Bill(s) or Charge(s)

                        //Generate Bills
                        //accountBillsSynced = await CreateDoctorsAndPayeesBills(serviceContext, ProcessId, caseDetails, _UserModalobj);

                        if (accountBillsSynced > 0)
                            result = (decimal)accountBillsSynced;

                        #endregion

                        #region LogQuickBooksSync

                        QuickBookLogCQB _QuickBookLogCQB = new QuickBookLogCQB();
                        _QuickBookLogCQB.AccessToken = access_token; //acces_token variable is defined in AppController as this controller is inherited from AppController
                        _QuickBookLogCQB.RecordsPushed = Convert.ToInt32(result);
                        _QuickBookLogCQB.SyncRequest = "AccountExports";
                        _QuickBookLogCQB.RealmId = realmId;
                        _QuickBookLogCQB.CreatedDate = DateTime.Now;
                        _QuickBookLogCQB.SyncUptoDate = DateTime.Now;
                        bool SyncDetails = QuickBooksManager.QuickBooksSyncLog(_QuickBookLogCQB);

                        #endregion
                        //Proces Status = Exported
                        status = await CaseManager.UpdateProcessStatus(ProcessId, 14, "", "", result.ToString(), 0);

                    }
                    //else
                    //return Json("-3", JsonRequestBehavior.AllowGet);//return View("Index", (object)"QBO API call Failed!");
                }
            }
            catch (Exception ex)
            {
                //return View("Index", (object)"QBO API calls Failed!");
                await CaseManager.DeleteProcessFromTemp(ProcessId);
                CustomLogger.LogError.GetLogger.Log(ex, "Error in ExportToQuickBooks Function. || Access Token = " + access_token + " || DateTime" + Convert.ToString(DateTime.Now));
                QuickBookLogCQB _QuickBookLogCQB = new QuickBookLogCQB();
                _QuickBookLogCQB.AccessToken = access_token; //acces_token variable is defined in AppController as this controller is inherited from AppController
                _QuickBookLogCQB.RecordsPushed = Convert.ToInt32(result);
                _QuickBookLogCQB.SyncRequest = "AccountExports";
                _QuickBookLogCQB.RealmId = realmId;
                _QuickBookLogCQB.CreatedDate = DateTime.Now;
                //_QuickBookLogCQB.SyncUptoDate = uptoDate;
                _QuickBookLogCQB.FailureReason = "Billing(), Reason: " + Convert.ToString(ex.InnerException.Message == null ? ex.Message : ex.InnerException.Message);
                bool SyncDetails = QuickBooksManager.QuickBooksSyncLog(_QuickBookLogCQB);
                //return Json("-2", JsonRequestBehavior.AllowGet);
                await CaseManager.UpdateProcessFailure(ProcessId, ex.InnerException.ToString(), 1);
            }
        }
        public async System.Threading.Tasks.Task AdddoctorsandPayeestoQBs(ServiceContext serviceContext, int ProcessId, UserModal _UserModalobj)
        {
            #region Get/Sync Vendor(s) or ServiceProvider(s)
            bool status = false;
            int accountPayablesSynced = 0;

            var doctorsToSync = await QuickBooksManager.GetSyncDoctors();

            if (doctorsToSync.Count > 0 && doctorsToSync != null)
            {
                accountPayablesSynced = await AddDoctorsListQB(serviceContext, doctorsToSync, DateTime.Now, _UserModalobj);
            }

            //Proces Status = Doctor Synced
            status = await CaseManager.UpdateProcessStatus(ProcessId, 12, "", "", "", 0);

            var payeesToSync = await QuickBooksManager.GetSyncPayees();

            if (payeesToSync.Count > 0 && payeesToSync != null)
            {
                accountPayablesSynced = await AddPayeeListQB(serviceContext, payeesToSync, DateTime.Now, _UserModalobj);
            }

            //Proces Status = Payee Synced
            status = await CaseManager.UpdateProcessStatus(ProcessId, 13, "", "", "", 0);
            #endregion
        }
        private async Task<int> AddDoctorsListQB(ServiceContext serviceContext, List<DoctorsAndPayees> vendors, DateTime uptoDate, UserModal _UserModalobj)
        {
            int count = 0;
            try
            {
                DataService dataService = new DataService(serviceContext);

                QueryService<Vendor> vendorQueryService = new QueryService<Vendor>(serviceContext);
                foreach (var vnd in vendors)
                {
                    count++;

                    Vendor vendor = new Vendor();

                    vendor.GivenName = vnd.GivenName;
                    vendor.MiddleName = vnd.MiddleName;
                    vendor.FamilyName = vnd.FamilyName;
                    vendor.Suffix = vnd.Suffix;
                    vendor.CompanyName = vnd.CompanyName;
                    vendor.DisplayName = vnd.DisplayName;
                    vendor.Active = true;
                    vendor.ActiveSpecified = true;

                    //Add contact details ie phone, fax, email and website details
                    if (vnd.PrimaryPhone.FreeFormNumber != null)
                    {
                        vendor.PrimaryPhone = new TelephoneNumber()
                        {
                            FreeFormNumber = vnd.PrimaryPhone.FreeFormNumber.Trim(),
                        };
                    }

                    vendor.BillAddr = new PhysicalAddress()
                    {
                        Line1 = vnd.BillAddr.Line1,
                        Country = vnd.BillAddr.Country
                    };

                    vendor.Fax = new TelephoneNumber()
                    {
                        FreeFormNumber = vnd.Fax.FreeFormNumber,
                    };

                    if (vnd.PrimaryEmailAddr.Address != null)
                    {
                        vendor.PrimaryEmailAddr = new EmailAddress()
                        {
                            Address = vnd.PrimaryEmailAddr.Address,
                        };
                    }

                    vendor.WebAddr = new WebSiteAddress()
                    {
                        URI = vnd.WebAddr.URI
                    };

                    try
                    {
                        System.Threading.Thread.Sleep(1300);

                        var vendorAdded = dataService.Add<Vendor>(vendor);
                        await QuickBooksManager.SaveDoctorsApiId(vnd.doctorOrPayeeId, Convert.ToString(vendorAdded.Id).Trim());
                    }
                    catch (Exception ex)
                    {
                        CustomLogger.LogError.GetLogger.Log(ex, "Error in CreateVendor Function. || Access Token = " + access_token + " || DateTime" + Convert.ToString(DateTime.Now));
                        QuickBookLogCQB _QuickBookLogCQB = new QuickBookLogCQB();
                        _QuickBookLogCQB.AccessToken = access_token; //acces_token variable is defined in AppController as this controller is inherited from AppController
                        _QuickBookLogCQB.RecordsPushed = Convert.ToInt32(count);
                        _QuickBookLogCQB.SyncRequest = "AccountPayables";
                        _QuickBookLogCQB.RealmId = Session["realmId"].ToString();
                        _QuickBookLogCQB.CreatedDate = DateTime.Now;
                        _QuickBookLogCQB.SyncUptoDate = DateTime.Now;
                        _QuickBookLogCQB.CreatedBy = _UserModalobj.UserId;
                        _QuickBookLogCQB.FailureReason = "CreateVendors() Inner, Reason: " + ex.InnerException.Message == null ? ex.Message : ex.InnerException.Message;
                        bool SyncDetails = await QuickBooksManager.QuickBooksSyncLogAsync(_QuickBookLogCQB);

                        continue;
                        // throw;
                    }

                    //return vendorAdded;
                }
                return count;
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in CreateVendor Function. || Access Token = " + access_token + " || DateTime" + Convert.ToString(DateTime.Now));
                QuickBookLogCQB _QuickBookLogCQB = new QuickBookLogCQB();
                _QuickBookLogCQB.AccessToken = access_token; //acces_token variable is defined in AppController as this controller is inherited from AppController
                _QuickBookLogCQB.RecordsPushed = Convert.ToInt32(count);
                _QuickBookLogCQB.SyncRequest = "AccountPayables";
                _QuickBookLogCQB.RealmId = Session["realmId"].ToString();
                _QuickBookLogCQB.CreatedDate = DateTime.Now;
                _QuickBookLogCQB.SyncUptoDate = DateTime.Now;
                _QuickBookLogCQB.CreatedBy = _UserModalobj.UserId;
                _QuickBookLogCQB.FailureReason = "CreateVendors(), Reason: " + Convert.ToString(ex.InnerException.Message == null ? ex.Message : ex.InnerException.Message);
                bool SyncDetails = await QuickBooksManager.QuickBooksSyncLogAsync(_QuickBookLogCQB);
                throw;
            }
        }

        private async Task<int> AddPayeeListQB(ServiceContext serviceContext, List<DoctorsAndPayees> vendors, DateTime uptoDate, UserModal _UserModalobj)
        {
            int count = 0;
            try
            {
                DataService dataService = new DataService(serviceContext);

                QueryService<Vendor> vendorQueryService = new QueryService<Vendor>(serviceContext);
                foreach (var vnd in vendors)
                {
                    count++;

                    Vendor vendor = new Vendor();

                    vendor.GivenName = vnd.GivenName.Trim();
                    vendor.MiddleName = vnd.MiddleName;
                    vendor.FamilyName = vnd.FamilyName;
                    vendor.Suffix = vnd.Suffix;
                    vendor.CompanyName = vnd.CompanyName;
                    vendor.DisplayName = vnd.DisplayName;

                    vendor.Active = true;
                    vendor.ActiveSpecified = true;

                    //Add contact details ie phone, fax, email and website details
                    if (vnd.PrimaryPhone.FreeFormNumber != null)
                    {
                        vendor.PrimaryPhone = new TelephoneNumber()
                        {
                            FreeFormNumber = vnd.PrimaryPhone.FreeFormNumber.Trim(),
                        };
                    }


                    vendor.BillAddr = new PhysicalAddress()
                    {
                        Line1 = vnd.BillAddr.Line1,
                        Country = vnd.BillAddr.Country
                    };

                    vendor.Fax = new TelephoneNumber()
                    {
                        FreeFormNumber = vnd.Fax.FreeFormNumber,
                    };

                    if (vnd.PrimaryEmailAddr.Address != null)
                    {
                        vendor.PrimaryEmailAddr = new EmailAddress()
                        {
                            Address = vnd.PrimaryEmailAddr.Address,
                        };
                    }

                    vendor.WebAddr = new WebSiteAddress()
                    {
                        URI = vnd.WebAddr.URI
                    };

                    try
                    {
                        System.Threading.Thread.Sleep(1300);

                        var vendorAdded = dataService.Add<Vendor>(vendor);
                        await QuickBooksManager.SavePayeesApiId(vnd.doctorOrPayeeId, Convert.ToString(vendorAdded.Id).Trim());
                    }
                    catch (Exception ex)
                    {
                        CustomLogger.LogError.GetLogger.Log(ex, "Error in CreateVendor Function. || Access Token = " + access_token + " || DateTime" + Convert.ToString(DateTime.Now));
                        QuickBookLogCQB _QuickBookLogCQB = new QuickBookLogCQB();
                        _QuickBookLogCQB.AccessToken = access_token; //acces_token variable is defined in AppController as this controller is inherited from AppController
                        _QuickBookLogCQB.RecordsPushed = Convert.ToInt32(count);
                        _QuickBookLogCQB.SyncRequest = "AccountPayables";
                        _QuickBookLogCQB.RealmId = Session["realmId"].ToString();
                        _QuickBookLogCQB.CreatedDate = DateTime.Now;
                        _QuickBookLogCQB.SyncUptoDate = uptoDate;
                        _QuickBookLogCQB.CreatedBy = _UserModalobj.CreatedBy;
                        _QuickBookLogCQB.FailureReason = "CreateVendors() Inner, Reason: " + ex.InnerException.Message == null ? ex.Message : ex.InnerException.Message;
                        bool SyncDetails = await QuickBooksManager.QuickBooksSyncLogAsync(_QuickBookLogCQB);

                        continue;
                        // throw;
                    }

                    //return vendorAdded;
                }
                return count;
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in CreateVendor Function. || Access Token = " + access_token + " || DateTime" + Convert.ToString(DateTime.Now));
                QuickBookLogCQB _QuickBookLogCQB = new QuickBookLogCQB();
                _QuickBookLogCQB.AccessToken = access_token; //acces_token variable is defined in AppController as this controller is inherited from AppController
                _QuickBookLogCQB.RecordsPushed = Convert.ToInt32(count);
                _QuickBookLogCQB.SyncRequest = "AccountPayables";
                _QuickBookLogCQB.RealmId = Session["realmId"].ToString();
                _QuickBookLogCQB.CreatedDate = DateTime.Now;
                _QuickBookLogCQB.SyncUptoDate = uptoDate;
                _QuickBookLogCQB.CreatedBy = _UserModalobj.CreatedBy;
                _QuickBookLogCQB.FailureReason = "CreateVendors(), Reason: " + Convert.ToString(ex.InnerException.Message == null ? ex.Message : ex.InnerException.Message);
                bool SyncDetails = await QuickBooksManager.QuickBooksSyncLogAsync(_QuickBookLogCQB);
                throw;
            }

        }
        private async Task<int> CreateDoctorsAndPayeesBills(ServiceContext serviceContext, int ProcessId, List<CaseDetailsModel> caseDetails, UserModal _UserModalobj)//List<CustomVendor> vendors)
        {
            int count = 0;
            try
            {
                CaseManager caseManager = new CaseManager();
                DataService dataService = new DataService(serviceContext);

                #region create liability account
                //Get a liability account. If not present create one
                QueryService<Account> accountQuerySvc = new QueryService<Account>(serviceContext);
                Account liabilityAccount = accountQuerySvc.ExecuteIdsQuery("SELECT * FROM Account WHERE AccountType='Accounts Payable' AND Classification='Liability'").FirstOrDefault();
                if (liabilityAccount == null)
                {
                    Account accountp = new Account();
                    String guid = Guid.NewGuid().ToString("N");
                    accountp.Name = "Name_" + guid;

                    // accountp.FullyQualifiedName = liabilityAccount.Name;

                    accountp.Classification = AccountClassificationEnum.Liability;
                    accountp.ClassificationSpecified = true;
                    accountp.AccountType = AccountTypeEnum.AccountsPayable;
                    accountp.AccountTypeSpecified = true;

                    accountp.CurrencyRef = new ReferenceType()
                    {
                        name = "United States Dollar",
                        Value = "USD"
                    };
                    liabilityAccount = dataService.Add<Account>(accountp);
                }
                #endregion

                #region create expense account
                //Get a Expense account. If not present create one
                Account expenseAccount = accountQuerySvc.ExecuteIdsQuery("SELECT * FROM Account WHERE AccountType='Expense' AND Classification='Expense'").FirstOrDefault();
                if (expenseAccount == null)
                {
                    Account accounte = new Account();
                    String guid = Guid.NewGuid().ToString("N");
                    accounte.Name = "Name_" + guid;

                    // accounte.FullyQualifiedName = expenseAccount.Name;

                    accounte.Classification = AccountClassificationEnum.Liability;
                    accounte.ClassificationSpecified = true;
                    accounte.AccountType = AccountTypeEnum.AccountsPayable;
                    accounte.AccountTypeSpecified = true;

                    accounte.CurrencyRef = new ReferenceType()
                    {
                        name = "United States Dollar",
                        Value = "USD"
                    };
                    expenseAccount = dataService.Add<Account>(accounte);
                }
                #endregion

                foreach (var caseDetail in caseDetails)
                {
                    var checkCaseSYNCED = QuickBooksManager.CheckIfCaseIsSyncedToQBs(caseDetail.caseIds);
                    if (checkCaseSYNCED == false)
                    {
                        var doctorOrPayeeApiId = QuickBooksManager.GetDoctorOrpayeeApiId(caseDetail.doctorIds, caseDetail.doctorLocationIds);

                        QueryService<Vendor> vendorQuerySvc = new QueryService<Vendor>(serviceContext);
                        var vendorInfo = vendorQuerySvc.ExecuteIdsQuery("Select * From vendor where Id = '" + doctorOrPayeeApiId + "'").FirstOrDefault();
                        #region create bill for the added vendor
                        if (vendorInfo != null)
                        {

                            //Create a bill and add a vendor reference
                            Bill bill = new Bill();

                            bill.DueDate = DateTime.UtcNow.Date;
                            bill.DueDateSpecified = true;
                            bill.PrivateNote = caseDetail.claimantName;
                            bill.VendorRef = new ReferenceType()
                            {

                                name = vendorInfo.DisplayName,
                                Value = vendorInfo.Id
                            };
                            bill.APAccountRef = new ReferenceType()
                            {

                                name = liabilityAccount.Name,
                                Value = liabilityAccount.Id
                            };
                            //Create a line for the bill
                            List<Line> lineList = new List<Line>();

                            Line line = new Line();
                            line.Description = "Charges for Claimant: " + caseDetail.claimantName;
                            line.Amount = caseDetail.caseVendorCharges;//new Decimal(100.00);
                            line.AmountSpecified = true;
                            line.DetailType = LineDetailTypeEnum.AccountBasedExpenseLineDetail;
                            line.DetailTypeSpecified = true;
                            line.Id = caseDetail.caseIds.ToString();

                            //Create an AccountBasedExpenseLineDetail
                            AccountBasedExpenseLineDetail detail = new AccountBasedExpenseLineDetail();
                            //detail.CustomerRef = new ReferenceType { name = customerInfo.DisplayName, Value = customerInfo.Id };
                            detail.AccountRef = new ReferenceType { name = expenseAccount.Name, Value = expenseAccount.Id };
                            detail.BillableStatus = BillableStatusEnum.NotBillable;
                            line.AnyIntuitObject = detail;

                            //return billadded;
                            lineList.Add(line);
                            bill.Line = lineList.ToArray();
                            try
                            {
                                var billadded = dataService.Add<Bill>(bill);
                                QuickBooksManager.SaveCaseChargesApiId(caseDetail.caseIds, billadded.Id.Trim(), caseDetail.caseVendorCharges);
                                count++;
                                QuickBooksManager.SaveCaseChargesHistory(caseDetail.caseIds, caseDetail.caseNumbers, caseDetail.caseVendorCharges, vendorInfo.DisplayName, _UserModalobj);
                                await caseManager.DeleteProcessCasesFromTemp(caseDetail.caseIds, ProcessId);
                            }
                            catch (Exception ex)
                            {
                                await caseManager.DeleteProcessCasesFromTemp(caseDetail.caseIds, ProcessId);
                                CustomLogger.LogError.GetLogger.Log(ex, "Error in CreateBill Function. || Access Token = " + access_token + " || DateTime" + Convert.ToString(DateTime.Now));
                                QuickBookLogCQB _QuickBookLogCQB = new QuickBookLogCQB();
                                _QuickBookLogCQB.AccessToken = access_token; //acces_token variable is defined in AppController as this controller is inherited from AppController
                                _QuickBookLogCQB.RecordsPushed = Convert.ToInt32(count);
                                _QuickBookLogCQB.SyncRequest = "AccountPayables";
                                _QuickBookLogCQB.RealmId = Session["realmId"].ToString();
                                _QuickBookLogCQB.CreatedDate = DateTime.Now;
                                _QuickBookLogCQB.SyncUptoDate = DateTime.Now;
                                _QuickBookLogCQB.FailureReason = "CreateBills() Inner, Reason: " + Convert.ToString(ex.InnerException.Message == null ? ex.Message : ex.InnerException.Message);
                                bool SyncDetails = QuickBooksManager.QuickBooksSyncLog(_QuickBookLogCQB);
                                continue;
                            }
                        }
                        else
                        {
                            await caseManager.DeleteProcessCasesFromTemp(caseDetail.caseIds, ProcessId);
                            await caseManager.UpdateDoctorsNotAvailableInQBs(doctorOrPayeeApiId);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CaseManager CaseManager = new CaseManager();
                await CaseManager.DeleteProcessFromTemp(ProcessId);
                CustomLogger.LogError.GetLogger.Log(ex, "Error in CreateBill Function" + "|| Access Token = " + access_token + " || DateTime" + Convert.ToString(DateTime.Now));
                QuickBookLogCQB _QuickBookLogCQB = new QuickBookLogCQB();
                _QuickBookLogCQB.AccessToken = access_token; //acces_token variable is defined in AppController as this controller is inherited from AppController
                _QuickBookLogCQB.RecordsPushed = Convert.ToInt32(count);
                _QuickBookLogCQB.SyncRequest = "AccountPayables";
                _QuickBookLogCQB.RealmId = Session["realmId"].ToString();
                _QuickBookLogCQB.CreatedDate = DateTime.Now;
                _QuickBookLogCQB.SyncUptoDate = DateTime.Now;
                _QuickBookLogCQB.FailureReason = "CreateBills() Outer, Reason: " + Convert.ToString(ex.InnerException.Message == null ? ex.Message : ex.InnerException.Message);
                bool SyncDetails = QuickBooksManager.QuickBooksSyncLog(_QuickBookLogCQB);
                //throw;
            }
            return count;
            #endregion
        }
        
    }
}