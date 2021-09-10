using Buisness.Services;
using Common;
using Common.Case.Models;
using Common.Master;
using DataAccess.CustomModal;
using DataAccess.DataModels;
using Newtonsoft.Json;
using SchedulingApp.Areas.Case.Models;
using SchedulingApp.CommonClass;
using SchedulingApp.Controllers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace SchedulingApp.Areas.Case.Controllers
{
    public class ReportController : BaseController
    {
        CaseManager CaseManager;
        ReportManager ReportManager;
        DoctorManager doctorManager;
        CommonManager commonManager;
        UserManager userManager;

        public ReportController()
        {
            CaseManager = new CaseManager();
            ReportManager = new ReportManager();
            doctorManager = new DoctorManager();
            commonManager = new CommonManager();
            userManager = new UserManager();
        }
        // GET: Case/Report
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public JsonResult GetDoctorDailyStatusReport()
        {

            var tempIMEReportFiltter = JsonConvert.DeserializeObject<List<IMEReportFiltter>>(Request.Params["IMEReportFiltter"]);
            IMEReportFiltter IMEReportFiltter = new IMEReportFiltter();
            IMEReportFiltter.For = tempIMEReportFiltter[0].For;
            IMEReportFiltter.FromIsValid = tempIMEReportFiltter[0].FromIsValid;
            IMEReportFiltter.DoctorId = tempIMEReportFiltter[0].DoctorId;
            IMEReportFiltter.LocationId = tempIMEReportFiltter[0].LocationId;
            IMEReportFiltter.Offset = tempIMEReportFiltter[0].Offset;
            IMEReportFiltter.DateOffset = tempIMEReportFiltter[0].DateOffset;
            IMEReportFiltter.Search = tempIMEReportFiltter[0].Search;
            IMEReportFiltter.SortOrder = tempIMEReportFiltter[0].SortOrder;
            IMEReportFiltter.SortBy = tempIMEReportFiltter[0].SortBy;
            IMEReportFiltter.For = Convert.ToDateTime(tempIMEReportFiltter[0].FromText);
            IMEReportFiltter.To = Convert.ToDateTime(tempIMEReportFiltter[0].ToText);

            List<IMEDataReport2Updated> _IMEDataReport2 = new List<IMEDataReport2Updated>();

            try
            {
                int recordsTotal = 0;
                int recordsFiltered = 0;
                var draw = Request.Form.GetValues("draw").FirstOrDefault();
                var start = Request.Form.GetValues("start").FirstOrDefault();
                var length = Request.Form.GetValues("length").FirstOrDefault();

                var sortColumn = Request.Form.GetValues("columns[" + Request.Form.GetValues("order[0][column]").FirstOrDefault() + "][name]").FirstOrDefault();
                var sortColumnDir = Request.Form.GetValues("order[0][dir]").FirstOrDefault();

                var searchvalue = Request["search[value]"] ?? "";

                _IMEDataReport2 = ReportManager.GetDoctorDailyStatusReport(IMEReportFiltter, draw, start, length, searchvalue, sortColumn, sortColumnDir);

                if (_IMEDataReport2 != null && _IMEDataReport2.Count != 0)
                {
                    recordsTotal = _IMEDataReport2[0].recordsTotal;
                    recordsFiltered = _IMEDataReport2[0].filterTotal;
                    _IMEDataReport2[0].Search = searchvalue;
                    _IMEDataReport2[0].SortBy = sortColumn;
                    _IMEDataReport2[0].SortOrder = sortColumnDir;
                }

                return Json(new { draw, recordsFiltered, recordsTotal, data = _IMEDataReport2 }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in Reports/GetDoctorDailyStatusReport Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                return Json("-2", JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public JsonResult ExportReportTo(IMEReportFiltter IMEReportFiltter)
        {
            int _return = 0;
            string TempaltePath = string.Empty;
            IMEReportFiltter.From = Convert.ToDateTime(IMEReportFiltter.FromText);
            IMEReportFiltter.For = Convert.ToDateTime(IMEReportFiltter.FromText);
            IMEReportFiltter.To = Convert.ToDateTime(IMEReportFiltter.ToText);

            var documentName = Guid.NewGuid().ToString();
            List<IMEDataReport2Updated> _IMEDataReport2 = new List<IMEDataReport2Updated>();
            try
            {


                _IMEDataReport2 = ReportManager.GetDoctorDailyStatusReport(IMEReportFiltter, "0", "0", "0", "", "AppointmentTimeString", "asc");

            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in Reports/GetIMEDataReportUpdated/GetDoctorDailyStatusReport Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                return Json("-1", JsonRequestBehavior.AllowGet);
            }
            try
            {
                if (Session[SessionEnum.loggedinuser.ToString()] != null)
                {
                    //AddDocumentModal.DocumentName = (AddDocumentModal.TemplateName ?? "") + "_" + Guid.NewGuid().ToString() + ".docx";                      

                    documentName = "DailyStatus" + "_" + documentName + ".docx";

                    TempaltePath = new DirectoryInfo(string.Format("{0}Documents\\Templates\\" + "Doctor Daily status.docx", Server.MapPath(@"\"))).ToString();

                    if (System.IO.File.Exists(TempaltePath.ToString()))
                    {
                        if (!System.IO.Directory.Exists(string.Format("{0}Documents\\Temp\\Reports\\", Server.MapPath(@"\"))))
                        {
                            System.IO.Directory.CreateDirectory(string.Format("{0}Documents\\Temp\\Reports\\", Server.MapPath(@"\")));
                        }
                        var originalDirectory = new DirectoryInfo(string.Format("{0}Documents\\Temp\\Reports\\", Server.MapPath(@"\")));
                        string pathString = System.IO.Path.Combine(originalDirectory.ToString(), /*"\\Reports\\"*/"");
                        bool isExists = System.IO.Directory.Exists(pathString);
                        if (!isExists)
                            System.IO.Directory.CreateDirectory(pathString);
                        var FilePath = new DirectoryInfo(string.Format("{0}Documents\\Temp\\Reports\\" + "File" + documentName, Server.MapPath(@"\"))).ToString();
                        if (!System.IO.File.Exists(FilePath))
                        {
                            if (!System.IO.Directory.Exists(string.Format("{0}Documents\\Temp\\", Server.MapPath(@"\"))))
                            {
                                System.IO.Directory.CreateDirectory(string.Format("{0}Documents\\Temp\\", Server.MapPath(@"\")));
                            }

                            var TempPath = new DirectoryInfo(string.Format("{0}Documents\\Temp\\Reports\\" + "Temp" + documentName, Server.MapPath(@"\"))).ToString();
                            System.IO.File.Copy(TempaltePath, TempPath);

                            var IsMergeSuccess = OpenXmlMailMerge.GetOpenXmlMailMerge.OpenXmlMergeForReports(TempPath, FilePath, _IMEDataReport2, IMEReportFiltter);


                            System.IO.File.Delete(TempPath);

                            if (IsMergeSuccess == false)
                            {
                                return Json("-1", JsonRequestBehavior.AllowGet);
                            }
                            return Json(new { filename = "File" + documentName });
                        }
                        else
                        {
                            Exception exception = new Exception("file does not exist: " + FilePath + "");
                            throw (exception);
                        }
                    }
                    else
                    {
                        ///file does not exist
                        _return = -1;
                    }
                }
                _return = 1;
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in Report/ExportReportTo Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                try
                {
                    if (ex != null)
                    {
                        CustomLogger.LogError.GetLogger.Log(ex, "Error in Reports/ExportReportTo Function.");
                        CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());                      
                    }
                }
                catch (Exception ex1)
                {
                    CustomLogger.LogError.GetLogger.Log(ex1, "Error in Reports/ExportReportTo writing error Function");
                    CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                }
                _return = -2;
            }
            return Json(_return, JsonRequestBehavior.AllowGet);
        }


    }
}