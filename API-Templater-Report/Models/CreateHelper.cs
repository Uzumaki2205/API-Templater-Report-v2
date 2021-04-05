using Newtonsoft.Json.Linq;
using NGS.Templater;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Web;

namespace API_Templater_Report.Models
{
    class CreateHelper
    {
        
    }
    public class InfoVuln
    {
        // private string rootPath = AppDomain.CurrentDomain.BaseDirectory;
        private string rootPath = HttpContext.Current.Server.MapPath($"~/");
        public string TimeStamp { get; set; }
        private dynamic infoObject { get; set; }
        private dynamic tableObject { get; set; }

        private String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyyMMddHHmmssffff");
        }

        private object ImageLoader(object value, string metadata)
        {
            //Plugin can be used to convert string into an Image type which Templater recognizes
            if (metadata == "from-resource" && value is string)
                return Image.FromFile(rootPath + $"/Static/{TimeStamp}/" + value.ToString());
            return value;
        }
        private object ImageMaxSize(object value, string metadata)
        {
            var bmp = value as Bitmap;
            if (metadata.StartsWith("maxSize(") && bmp != null)
            {
                var parts = metadata.Substring(8, metadata.Length - 9).Split(',');
                var maxWidth = int.Parse(parts[0].Trim()) * 28;
                var maxHeight = int.Parse(parts[parts.Length - 1].Trim()) * 28;
                if (bmp.Width > 0 && maxWidth > 0 && bmp.Width > maxWidth || bmp.Height > 0 && maxHeight > 0 && bmp.Height > maxHeight)
                {
                    var widthScale = 1f * bmp.Width / maxWidth;
                    var heightScale = 1f * bmp.Height / maxHeight;
                    var scale = Math.Max(widthScale, heightScale);
                    //Before passing image for processing it can be manipulated via Templater plugins
                    bmp.SetResolution(bmp.HorizontalResolution * scale, bmp.VerticalResolution * scale);
                }
            }
            return value;
        }
        public void ProcessDocx(string nameTemplate, JObject jObject)
        {
            TimeStamp = GetTimestamp(DateTime.Now);
            CreateFolder();

            infoObject = new List<Object>();
            tableObject = new List<Object>();

            Dictionary<string, string> obj_Property = new Dictionary<string, string>();
            List<Dictionary<string, string>> obj_Array = new List<Dictionary<string, string>>();
            List<Dictionary<string, Object>> temp = new List<Dictionary<string, Object>>(); // Xử lý bảng

            File.Copy(HttpContext.Current.Server.MapPath($"~/Template/{nameTemplate}"), 
                HttpContext.Current.Server.MapPath($"~/Renders/{TimeStamp}.Report.docx"), true);

            var factory = Configuration.Builder
            .Include(ImageLoader)   //setup image loading via from-resource metadata
            .Include(ImageMaxSize)  //setup image resizing via maxSize(X, Y) metadata
            .Include(ColorConverter)    //setup image from color converter
            .Build();
            //using (StreamReader r = new StreamReader(json_http))
            //{
            //string json = r.ReadToEnd();
            //JObject jObject = JObject.Parse(json_http);

            using (var doc = factory.Open(HttpContext.Current.Server.MapPath($"~/Renders/{TimeStamp}.Report.docx")))
            {
                foreach (JProperty property in jObject.Properties())
                {
                    foreach (var item in property)
                    {
                        if (item.Type == JTokenType.String)
                        {
                            obj_Property.Add(property.Name, property.Value.ToString());
                        }
                        else if (item.Type == JTokenType.Object)
                        {
                            if (property.Name.ToLower().Equals("chart"))
                            {
                                Dictionary<string, object>[] chart = new Dictionary<string, object>[4];
                                int i = 0;
                                foreach (JProperty itemChart in item)
                                {
                                    chart[i] = new Dictionary<string, object>() { { "name", itemChart.Name }, { "value", itemChart.Value } };
                                    i++;
                                }
                                doc.Process(new[] { new { pie = chart } });
                            }
                            else // If into JObject have JArray
                            {
                                foreach (var objArrayProperty in item)
                                {
                                    foreach (var objArray in objArrayProperty)
                                    {
                                        temp = ProcessTable(objArray);
                                    }
                                }
                                doc.Process(temp);
                            }
                        }
                        else if (item.Type == JTokenType.Array)
                        {
                            temp = ProcessTable(item);

                            infoObject.Add(obj_Property);

                            doc.Process(infoObject);
                            doc.Process(temp); //bảng
                        }
                    }
                }
            }

                DeleteFolderImage();
                //Process.Start(new ProcessStartInfo($"{TimeStamp}.docx") { UseShellExecute = true });
            //}
        }

        private List<Dictionary<string, Object>> ProcessTable(JToken array)
        {
            List<Dictionary<string, Object>> Dictionary_Looping = new List<Dictionary<string, Object>>();

            foreach (JToken item in array)
            {
                // Dictionary <string -> Key, Object -> Can is String type or List<Dictionary<string, string>> Type>
                Dictionary<string, Object> dic = new Dictionary<string, Object>();
                Dictionary<string, Object> dicObj = new Dictionary<string, Object>();
                foreach (JProperty subItem in item)
                {
                    if (subItem.Parent.Type == JTokenType.Object)
                    {
                        foreach (var arr2 in subItem)
                        {
                            if (arr2.Type == JTokenType.Array)
                            {
                                List<Dictionary<string, Object>> lstArr2 = new List<Dictionary<string, Object>>(); // List of array2 json
                                foreach (var lstarray in arr2)
                                {
                                    Dictionary<string, Object> array2p = new Dictionary<string, Object>();
                                    foreach (JProperty array2p2 in lstarray)
                                    {
                                        if (array2p2.Name.StartsWith("color-"))
                                            array2p.Add("Color", ProcessColor(array2p2));

                                        array2p.Add(array2p2.Name, array2p2.Value.ToString());
                                    }
                                    lstArr2.Add(array2p);
                                }
                                dic.Add(subItem.Name, lstArr2);
                            }
                            else
                            {
                                if (subItem.Name.StartsWith("color-"))
                                    dicObj.Add("Color", ProcessColor(subItem));

                                if (subItem.Name.StartsWith("image-"))
                                {
                                    string valueImage = ProcessImage(subItem);
                                    dicObj.Add(subItem.Name, valueImage + ".jpg");
                                }
                                else dicObj.Add(subItem.Name, subItem.Value.ToString());
                            }
                        }
                    }
                    //else if (subItem.Type == JTokenType.Array)
                    //{
                    //    foreach (var tempArr in subItem)
                    //    {
                    //        // If have array into array -> move array2 to List<string, string>
                    //        // Then add List<string, string> to Dictionary<string, Dictionary<string,string>>
                    //        if (tempArr.Type == JTokenType.Array)
                    //        {
                    //            List<Dictionary<string, Object>> lstArr2 = new List<Dictionary<string, Object>>(); // List of array2 json
                    //            foreach (JToken itemArray2p in tempArr)
                    //            {
                    //                Dictionary<string, Object> array2p = new Dictionary<string, Object>();
                    //                foreach (JProperty subItemArray2p in itemArray2p)
                    //                {
                    //                    array2p.Add(subItemArray2p.Name, subItemArray2p.Value.ToString());

                    //                    if (subItemArray2p.Name.StartsWith("color-"))
                    //                    {
                    //                        array2p.Add("Color", ProcessColor(subItemArray2p));
                    //                    }   
                    //                }
                    //                // Create List Dictionary of array into array :)))
                    //                lstArr2.Add(array2p);
                    //            }
                    //            dic.Add(subItem.Name, lstArr2);
                    //        }
                    //        else
                    //        {
                    //            if (subItem.Name.StartsWith("image-"))
                    //            {
                    //                string valueImage = ProcessImage(subItem);
                    //                dic.Add(subItem.Name, valueImage + ".jpg");
                    //            }
                    //            else
                    //            {
                    //                dic.Add(subItem.Name, subItem.Value.ToString());
                    //            }
                    //        }
                    //    }
                    //}

                }

                dic.Add(array.Path, dicObj);
                Dictionary_Looping.Add(dic);
            }

            return Dictionary_Looping;
        }

        //private Dictionary<string, object>[] ProgressChart(int critical, int high, int medium, int low)
        //{
        //    var pie1 = new Dictionary<string, object>[4];
        //    pie1[0] = new Dictionary<string, object>() { { "name", "Critical" }, { "value", critical } };
        //    pie1[1] = new Dictionary<string, object>() { { "name", "High" }, { "value", high } };
        //    pie1[2] = new Dictionary<string, object>() { { "name", "Medium" }, { "value", medium } };
        //    pie1[3] = new Dictionary<string, object>() { { "name", "Low" }, { "value", low } };

        //    return pie1;
        //}

        private Color ProcessColor(JProperty property)
        {
            Color color = Color.Transparent;

            if (property.Value.ToString().ToLower().Equals("critical") || property.Value.ToString().ToLower().Equals("nguy hiểm"))
                color = Color.DeepPink;
            else if (property.Value.ToString().ToLower().Equals("high") || property.Value.ToString().ToLower().Equals("cao"))
                color = Color.Red;
            else if (property.Value.ToString().ToLower().Equals("medium") || property.Value.ToString().ToLower().Equals("trung bình"))
                color = Color.Orange;
            else if (property.Value.ToString().ToLower().Equals("low") || property.Value.ToString().ToLower().Equals("thấp"))
                color = Color.Yellow;

            return color;
        }

        private string ProcessImage(JProperty image)
        {
            Random r = new Random();
            string temp = r.Next(500).ToString();
            Base64ToImage(image.Value.ToString(), temp);

            return temp;
        }

        //Convert Base64 to Image
        private void Base64ToImage(string base64, string name)
        {
            // Bitmap bm2;
            if (base64 == null || base64 == string.Empty)
                base64 = "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";
            byte[] base64Obj = null;
            try
            {
                base64Obj = Convert.FromBase64String(base64);
            }
            catch (Exception)
            {
                base64 = "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";
                base64Obj = Convert.FromBase64String(base64);
            }

            using (MemoryStream ms = new MemoryStream(base64Obj))
            {
                Image img = Image.FromStream(ms);
                img.Save(HttpContext.Current.Server.MapPath($"~/Static/{TimeStamp}/{name}.jpg"));
                ms.Close();
            }
        }


        //Create FOlder if not exist
        private void CreateFolder()
        {
            if (!Directory.Exists($"{rootPath}/Renders"))
                Directory.CreateDirectory($"{rootPath}/Renders");
            //if (!Directory.Exists($"{rootPath}/Static"))
            //    Directory.CreateDirectory($"{rootPath}/Static");
            if (!Directory.Exists($"{rootPath}/{TimeStamp}"))
                Directory.CreateDirectory($"{rootPath}/Static/{TimeStamp}");
        }
        //Delete Folder Image temp after insert into word
        private void DeleteFolderImage()
        {
            DirectoryInfo dir = new DirectoryInfo($"{rootPath}/Static/{TimeStamp}");

            if (dir.Exists)
            {
                try
                {
                    dir.Delete(true);
                }
                catch (Exception)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    dir.Delete(true);
                }
            }
        }

        private object ColorConverter(object value, string tag, string[] metadata)
        {
            if (value is Color)
            {
                var col = (Color)value;
                var fillValue = col.R.ToString("X2") + col.G.ToString("X2") + col.B.ToString("X2");
                return System.Xml.Linq.XElement.Parse(@"
                <w:tc xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                 <w:tcPr>
                  <w:shd w:val=""clear"" w:color=""auto"" w:fill=""" + fillValue + @""" />
                 </w:tcPr>
                </w:tc>");
            }
            return value;
        }

    }
}
