using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Common.Case.Models;
using System.IO;
using Buisness.Services;
using Common;
using SchedulingApp.Controllers;
using SchedulingApp.CommonClass;
using Common.Master;
using Common.Case.Template;
using Spire.Doc;
using System.Text;
using Common.QuickBooks;
using DocumentSp = Spire.Doc.Document;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Reporting;
using System.Drawing.Printing;
using System.Drawing;
using Spire.Pdf;
using WordToPDF;
//using Syncfusion.EJ2.DocumentEditor;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Document = Spire.Doc.Document;
//using OpenXmlPowerTools;
using DocumentFormat.OpenXml;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using System.Windows.Media.Imaging;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Net.Http;
using System.Configuration;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.WindowsAzure.Storage.Auth;
using System.Net;
using Syncfusion.EJ2.DocumentEditor;
using System.Globalization;
//using iTextSharp.text.pdf;
//using System.Text.RegularExpressions;
//using Microsoft.Office.Interop.Word;
namespace SchedulingApp.Areas.Case.Controllers
{
    public class DocmentsController : BaseController
    {

        CommonManager commonManager;
        DocumentManager DocumentManager;
        ServiceManager serviceManager;

        private static string BlobAccountUrl = ConfigurationManager.AppSettings["BlobAccountUrl"];
        public DocmentsController()
        {
            commonManager = new CommonManager();
            DocumentManager = new DocumentManager();
            serviceManager = new ServiceManager();
        }
        // GET: Case/Docments
        public ActionResult Index()
        {
            return View();
        }

        
        public async Task<JsonResult> CallChartPrepWebApi(AddDocumentModal AddDocumentModals)
        {
            bool success = true;
            try
            {
                HttpClient client = new HttpClient();

                UserModal _UserModalobj = (UserModal)Session[SessionEnum.loggedinuser.ToString()];
                var DocPath = Server.MapPath(@"\");
                AddDocumentModals.LocalServerPath = DocPath;

                AddDocumentModals.CurrentUserEmail = _UserModalobj.Email;
                AddDocumentModals.CurrentUserName = _UserModalobj.CreatedBy;
                AddDocumentModals.CurrentUserId = _UserModalobj.UserId;
                AddDocumentModals.LstPermissionsId = _UserModalobj.LstPermissionsId;
                AddDocumentModals.LstPermissions = _UserModalobj.LstPermissions;

                var jsonobj = JsonConvert.SerializeObject(AddDocumentModals);
                var stringContent = new StringContent(jsonobj, Encoding.UTF8, "application/json");

                //API Hit
                string APIUrl = ConfigurationManager.AppSettings["APIUrlDocments"];
                var response = client.PostAsync(APIUrl, stringContent);

                return Json("1", JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in Documents/CallChartPrepWebApi writing error Function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                success = false;
                return Json("-1", JsonRequestBehavior.AllowGet);
            }

        }

    }
}