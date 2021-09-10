using Common.Case.Models;
using DocumentWebApp;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using System.Web.Http.Results;

namespace DocumentsWebAPI.Controllers
{
    public class CustomAPIController : ApiController
    {
        // GET api/<controller>
        public IEnumerable<string> Get()
        {
            return new string[] { "Anjali", "Sumit" };
        }

        // GET api/<controller>/5
        public string Get(int id)
        {
            DocumentDownload documentDownload = new DocumentDownload();
            return documentDownload.demo();
        }

        public async Task Post(AddDocumentModal AddDocumentModals)
        {

            bool success = true;

            try
            {
                string url = Request.RequestUri.AbsoluteUri;
                CustomLogger.LogError.GetLogger.Log(new Exception(""), "Post-1 url=" + url);

                //return AddDocumentModals.BatchCaseId.ToString();
                DocumentDownload documentDownload = new DocumentDownload();
                var DocPath = HttpContext.Current.Server.MapPath(@"\");
                AddDocumentModals.LocalServerPath = DocPath;
                var str = documentDownload.GenerateTemplates(AddDocumentModals);

                CustomLogger.LogError.GetLogger.Log(new Exception(""), "Post-DocPath=" + DocPath + " url=" + url);
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(new Exception(""), "Api Throws Exception");
            }

        }
    }
}
