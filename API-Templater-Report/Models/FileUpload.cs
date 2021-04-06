using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace API_Templater_Report.Models
{

    /// <summary>
    /// Define File Upload
    /// </summary>
    public class FileUpload
    {
        public string FileId { get; set; }
        public string FileName { get; set; }
        public string FileUrl { get; set; }
        public string JsonName { get; set; }
        public string JsonUrl { get; set; }
        public IEnumerable<FileUpload> FileList { get; set; }
    }
}