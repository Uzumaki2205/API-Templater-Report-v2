using API_Templater_Report.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace API_Templater_Report.Controllers
{
    /// <summary>
    /// API TEMPLATE REPORT
    /// </summary>
    public class FillDocxController : ApiController
    {
        // GET: api/FillDocx
        /// <summary>
        /// ReadMe
        /// </summary>
        [HttpGet]
        public IEnumerable<string> Readme()
        {
            return new string[] 
            { 
                "POST to Url: /api/filldocx with JSON format to proccess file .DOCX",
                "GET /api/Download/start?fileName={fileName} to Download"
            };
        }

        /// <summary>
        /// Check File Exists
        /// </summary>
        /// <param name="fileName">Name of generated file</param>
        [HttpGet]
        public HttpResponseMessage IsExistFile(string fileName)
        {
            var filepath = HttpContext.Current.Server.MapPath($"~/Renders/{fileName}");
            Uri path = new Uri($"/api/Filldocx/Download/start?fileName={fileName}", UriKind.Relative);

            if (File.Exists(filepath))
            {
                var response = new HttpResponseMessage(HttpStatusCode.OK);
                response.Content = new StringContent($"<a href='{path}'>{path}</a>");
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
                return response;
            }
            else return new HttpResponseMessage(HttpStatusCode.NotFound);
        }

        InfoVuln helper = new InfoVuln();

        /// <summary>
        /// Download File
        /// </summary>
        /// <param name="fileName">Name of generated file</param>
        [HttpGet]
        public HttpResponseMessage Download(string fileName)
        {
            string filePath = HttpContext.Current.Server.MapPath($"~/Renders/{fileName}");

            if (File.Exists(filePath))
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);
                byte[] bytes = File.ReadAllBytes(filePath);
                response.Content = new ByteArrayContent(bytes);
                //Set the Response Content Length.
                response.Content.Headers.ContentLength = bytes.LongLength;

                //Set the Content Disposition Header Value and FileName.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                response.Content.Headers.ContentDisposition.FileName = fileName;
                //Set the File Content Type.
                response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(fileName));
                return response;
            }
            else return new HttpResponseMessage(HttpStatusCode.BadRequest);
        }

        // POST api/filldocx/Generate
        /// <summary>
        /// Generate Template
        /// </summary>
        /// <param name="nameTemplate">Template Name</param>
        /// <param name="json">JSON OBJECT</param>
        [HttpPost]
        public HttpResponseMessage Generate(string nameTemplate, [FromBody] JObject json)
        {
            try
            {
                helper.ProcessDocx(nameTemplate, json);
                return IsExistFile($"{helper.TimeStamp}.Report.docx");
            }
            catch (Exception)
            {
                //Console.WriteLine(ex.Message);
                return new HttpResponseMessage() { StatusCode = HttpStatusCode.NotFound };
            }
        }
    }
}
