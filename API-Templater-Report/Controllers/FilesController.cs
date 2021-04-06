using API_Templater_Report.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace API_Templater_Report.Controllers
{
    public class FilesController : Controller
    {
        // GET: Files
        //public ActionResult Index()
        //{
        //    return View();
        //}
        private string conString = @"Data Source=DESKTOP-SBRAGR1\SQLEXPRESS;Initial Catalog=Api-template-report;Integrated Security=True";
        public ActionResult Index(FileUpload model)
        {
            List<FileUpload> list = new List<FileUpload>();
            DataTable dtFiles = GetFileDetails();
            foreach (DataRow dr in dtFiles.Rows)
            {
                list.Add(new FileUpload
                {
                    FileId = @dr["Id"].ToString(),
                    FileName = @dr["FILENAME"].ToString(),
                    FileUrl = @dr["FILEURL"].ToString()
                    // Missing 2 configuration for JSON
                });
            }
            model.FileList = list;
            return View(model);
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
        public ActionResult Index(HttpPostedFileBase files)
        {
            string ext = Path.GetExtension(files.FileName);

            if (ext == ".doc" || ext == ".docx")
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
                        return RedirectToAction("Index", "Files");
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

            return RedirectToAction("Index", "Files");
        }

        // Save File After Uploaded
        private bool SaveFile(FileUpload model)
        {
            string strQryCheck = "SELECT COUNT(*) FROM [dbo].tblFileDetails where FILENAME= '" + model.FileName + "'";
            SqlConnection con = new SqlConnection(conString);
            con.Open();
            
            using (SqlCommand commandCheck = new SqlCommand(strQryCheck, con))
            {
                int numResultCheck = (int)commandCheck.ExecuteScalar();
                if (numResultCheck > 0)
                {
                    model.FileName = $"{DateTime.Now.ToString("yyyyMMddHHmmssffff")}.{model.FileName}";
                    TempData["AlertMessage"] = $"File changed to {model.FileName} !!";
                }  
            }
            
            string strQry = "INSERT INTO tblFileDetails (FileName, FileUrl, JsonName, JsonUrl) VALUES('" + model.FileName + "','" + model.FileUrl 
                + "','" + model.JsonName + "','" + model.JsonUrl + "')";
            SqlCommand command = new SqlCommand(strQry, con);
            int numResult = command.ExecuteNonQuery();
            con.Close();
            if (numResult > 0)
                return true;
            else
                return false;
        }

        public ActionResult DownloadFile(string filePath)
        {
            string fullName = Server.MapPath("~" + filePath);

            byte[] fileBytes = GetFile(fullName);
            return File(
                fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, filePath);
        }

        private byte[] GetFile(string s)
        {
            FileStream fs = System.IO.File.OpenRead(s);
            byte[] data = new byte[fs.Length];
            int br = fs.Read(data, 0, data.Length);
            if (br != fs.Length)
                throw new IOException(s);
            return data;
        }
    }
}