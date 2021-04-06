using API_Templater_Report.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Web;

using System.Web.Mvc;

namespace API_Templater_Report.Controllers
{
    public class TemplateController : Controller
    {
        private string conString = @"Data Source=DESKTOP-SBRAGR1\SQLEXPRESS;Initial Catalog=Api-template-report;Integrated Security=True";
        // GET: Template
        [HttpGet]
        public ActionResult Index(FileUpload model)
        {
            List<FileUpload> list = new List<FileUpload>();
            DataTable dtFiles = GetFileDetails();
            foreach (DataRow dr in dtFiles.Rows)
            {
                list.Add(new FileUpload
                {
                    //FileId = @dr["Id"].ToString(),
                    FileName = @dr["FILENAME"].ToString(),
                    FileUrl = @dr["FILEURL"].ToString()
                });
            }
            model.FileList = list;
            return View(list);
        }

        [HttpPost]
        public ActionResult UploadFile(HttpPostedFileBase files)
        {
            string ext = Path.GetExtension(files.FileName);

            if (ext == ".json")
            {
                FileUpload model = new FileUpload();
                List<FileUpload> list = new List<FileUpload>();
                DataTable dtFiles = GetFileDetails();
                foreach (DataRow dr in dtFiles.Rows)
                {
                    list.Add(new FileUpload
                    {
                        FileId = @dr["Id"].ToString(),
                        FileName = @dr["FILENAME"].ToString(),
                        FileUrl = @dr["FILEURL"].ToString(),
                        JsonName = @dr["JSONNAME"].ToString(),
                        JsonUrl = @dr["JSONURL"].ToString()
                    });
                }
                model.FileList = list;

                if (files != null)
                {
                    //var Extension = Path.GetExtension(files.FileName);
                    var timeStamp = InfoVuln.GetInstance().TimeStamp;
                    if (!Directory.Exists(Server.MapPath($"~/UploadedFiles/{timeStamp}")))
                        Directory.CreateDirectory(Server.MapPath($"~/UploadedFiles/{timeStamp}"));

                    string path = Path.Combine(Server.MapPath($"~/UploadedFiles/{timeStamp}"), files.FileName);
                    model.FileUrl = Url.Content(Path.Combine($"~/UploadedFiles/{timeStamp}/", files.FileName));

                    model.FileName = files.FileName;

                    if (SaveFile(model))
                    {
                        files.SaveAs(path);
                        TempData["AlertMessage"] = "Uploaded Successfully !!";
                        return RedirectToAction("Index", "Template");
                    }
                    else
                    {
                        ModelState.AddModelError("", "Error In Add File. Please Try Again !!!");
                    }
                }
                else
                {
                    ModelState.AddModelError("", "Please Choose Correct File Type !!");
                    return View(model);
                }
            }

            return RedirectToAction("Index", "Template");
        }

        // Save File After Uploaded
        private bool SaveFile(FileUpload model)
        {
            string strQryCheck = "SELECT COUNT(*) FROM [dbo].tblFileDetails where JSONNAME= '" + model.JsonName + "'";
            SqlConnection con = new SqlConnection(conString);
            con.Open();

            using (SqlCommand commandCheck = new SqlCommand(strQryCheck, con))
            {
                int numResultCheck = (int)commandCheck.ExecuteScalar();
                if (numResultCheck > 0)
                {
                    model.JsonName = $"{DateTime.Now.ToString("yyyyMMddHHmmssffff")}.{model.JsonName}";
                    TempData["AlertMessage"] = $"File changed to {model.JsonName} !!";
                }
            }

            //string strQry = "INSERT INTO tblFileDetails (FileName, FileUrl, JsonName, JsonUrl) VALUES('" + model.FileName + "','" + model.FileUrl
            //    + "','" + model.JsonName + "','" + model.JsonUrl + "')";
            string strQry = $"UPDATE tblFileDetails SET JsonName='{model.JsonName}', JsonUrl='{model.JsonUrl}' WHERE FILENAME='{model.FileName}'";
            SqlCommand command = new SqlCommand(strQry, con);
            int numResult = command.ExecuteNonQuery();
            con.Close();
            if (numResult > 0)
                return true;
            else
                return false;
        }

        private DataTable GetFileDetails()
        {
            DataTable dtData = new DataTable();
            SqlConnection con = new SqlConnection(conString);
            con.Open();
            SqlCommand command = new SqlCommand("Select * From tblFileDetails", con);
            SqlDataAdapter da = new SqlDataAdapter(command);
            da.Fill(dtData);
            con.Close();
            return dtData;
        }

        [HttpPost]
        public ContentResult Generate([System.Web.Http.FromBody] string nameTemplate, string json)
        {

            FillDocxController fillDocx = new FillDocxController();
            InfoVuln.GetInstance().ProcessDocx(nameTemplate, json);

            var timeStamp = InfoVuln.GetInstance().TimeStamp;
            if (fillDocx.IsExistFile(timeStamp + ".Report.docx").IsSuccessStatusCode)
                return base.Content($"<a href='/api/filldocx/download?filename={timeStamp}.Report.docx'>{timeStamp}.Report.docx</a>", "text/html");
            //$"/api/filldocx/download?filename={helper.TimeStamp}.Report.docx";
            else return base.Content("Not Found File");

            //string fullName = Server.MapPath("~/Render/" + helper.TimeStamp + ".Report.docx");

                //byte[] fileBytes = GetFile(fullName);
                //return File(
                //    fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, Server.MapPath("~/Render/" + helper.TimeStamp + ".Report.docx"));



                //if (fillDocx.IsExistFile(helper.TimeStamp + ".Report.docx").StatusCode == System.Net.HttpStatusCode.OK)
                //    return fillDocx.Download(helper.TimeStamp + ".Report.docx");
                //else return new HttpResponseMessage(System.Net.HttpStatusCode.NotFound);
        }
        //byte[] GetFile(string s)
        //{
        //    FileStream fs = System.IO.File.OpenRead(s);
        //    byte[] data = new byte[fs.Length];
        //    int br = fs.Read(data, 0, data.Length);
        //    if (br != fs.Length)
        //        throw new IOException(s);
        //    return data;
        //}

        //public ActionResult DownloadFile(string filePath)
        //{
        //    string fullName = Server.MapPath("~" + filePath);

        //    byte[] fileBytes = GetFile(fullName);
        //    return File(
        //        fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, filePath);
        //}
    }
}