using System;
using System.Net;
using Ionic.Zip;
using System.IO;
using System.Xml.Serialization;
using System.Text;
using System.Xml;
using System.Collections.Generic;

namespace Учебное
{
    class Program
    {
        private class CookieAwareWebClient : WebClient
        {
            public CookieAwareWebClient()
                : this(new CookieContainer())
            { }
            public CookieAwareWebClient(CookieContainer c)
            {
                this.CookieContainer = c;
            }
            public CookieContainer CookieContainer { get; set; }

            protected override WebRequest GetWebRequest(Uri address)
            {
                WebRequest request = base.GetWebRequest(address);

                var castRequest = request as HttpWebRequest;
                if (castRequest != null)
                {
                    castRequest.CookieContainer = this.CookieContainer;
                }

                return request;
            }
        }

        static  void Main(string[] args)
        {
            //Working part that gets cookie and html documetn from web
            //var cookies = new CookieContainer();
            //var request = (HttpWebRequest)WebRequest.Create("http://bo-otchet.1gl.ru/Account/Login");
            //request.CookieContainer = cookies;
            //request.Method = "POST";
            //request.ContentType = "application/x-www-form-urlencoded";
            //try
            //{
            //    using (var requestStream = request.GetRequestStream())
            //    using (var writer = new StreamWriter(requestStream))
            //    {
            //        writer.Write("Name=teleskopsoft&Password=118&returnUrl=");
            //    }

            //    using (var responseStream = request.GetResponse())
            //    cookies = request.CookieContainer;

            //    //Getting data from web by WebClient
            //    using (var client = new CookieAwareWebClient(cookies))
            //    {
            //        client.DownloadFile("http://bo-otchet.1gl.ru/Registration/RepostRegistrationExcel?Caption=%D0%A1%20%D0%B2%D1%8B%D0%BF%D1%83%D1%89%D0%B5%D0%BD%D0%BD%D0%BE%D0%B9%20%D0%9A%D0%AD%D0%9F&ShowSearchForm=False&PageId=0&PageSize=50&SuppressFlags=0&ViewId=HaveCertificate&PartnerId=118", "Test.xlsx");

            //        using (ZipFile zip = ZipFile.Read("Test.xlsx"))
            //        {
            //            if (zip == null) Console.WriteLine("Null");
            //            else
            //            {
            //                foreach (ZipEntry e in zip)
            //                {
            //                    e.Extract(ExtractExistingFileAction.OverwriteSilently);
            //                }
            //            }
            //        }
            //    }
                try
                {
                    XmlDocument _xml = new XmlDocument();
                    try
                    {
                        _xml.Load("xl\\sharedStrings.xml");
                    }
                    catch (System.IO.FileNotFoundException ex)
                    {
                        Console.WriteLine($"File not found!{Environment.NewLine}Error: {ex.ToString()}");
                    }
                    string str = _xml.OuterXml;
                    str = str.Replace("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"800\" uniqueCount=\"500\"", "");
                    XmlSerializer xml = new XmlSerializer(typeof(sst));
                    sst _sst = xml.Deserialize(new MemoryStream(Encoding.UTF8.GetBytes(str))) as sst;

                    try
                    {
                        _xml.Load("xl\\worksheets\\sheet1.xml");
                    }
                    catch (System.IO.FileNotFoundException ex)
                    {
                        Console.WriteLine($"File njt found!{Environment.NewLine}Error: {ex.ToString()}");
                    }
                    str = _xml.OuterXml;
                    str = str.Replace("xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"", "")
                             .Replace("xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"", "")
                             .Replace("x14ac:dyDescent=\"0.25\"", "")
                             .Replace("mc:Ignorable=\"x14ac\"", "")
                             .Replace("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "")
                             .Replace("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"", "");
                             
                    xml = new XmlSerializer(typeof(worksheet));
                    var _worksheet = xml.Deserialize(new MemoryStream(Encoding.UTF8.GetBytes(str))) as worksheet;
                    var _data = new List<EOData>();
                    EOData _temp = new EOData();

                    for (int j = 1; j <= _worksheet.Rows.GetLength(0)-1; j++)
                    {
                        _temp = new EOData();
                        _temp.ID = _worksheet.Rows[j].FilledCells[0].Value;
                        _temp.Long_ID = _worksheet.Rows[j].FilledCells[1].Value;
                        _temp.Name = _worksheet.Rows[j].FilledCells[2].Value;
                        _temp.INN = _worksheet.Rows[j].FilledCells[3].Value;
                        _temp.Type = _worksheet.Rows[j].FilledCells[15].Value;
                        _temp.Agent = _worksheet.Rows[j].FilledCells[4].Value;
                        _temp.Accountent = _worksheet.Rows[j].FilledCells[5].Value;
                        _temp.Email = _worksheet.Rows[j].FilledCells[6].Value;
                        _temp.Mobile = _worksheet.Rows[j].FilledCells[7].Value;
                        _temp.ES = _worksheet.Rows[j].FilledCells[8].Value;
                        _temp.PFR = _worksheet.Rows[j].FilledCells[9].Value;
                        _temp.Notification = _worksheet.Rows[j].FilledCells[10].Value;
                        _temp.Partner = _worksheet.Rows[j].FilledCells[11].Value;
                        _temp.LastUpdate = _worksheet.Rows[j].FilledCells[12].Value;
                        _temp.DateStart = _worksheet.Rows[j].FilledCells[13].Value;
                        _temp.DateEnd = _worksheet.Rows[j].FilledCells[14].Value;            
                        _data.Add(_temp);
                    }
                    _temp = null;
                    foreach (var item in _data)
                    {
                        Console.WriteLine(item.Print());
                    }


                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            //}
            //catch (WebException ex)
            //{
            //    if (ex.Status == WebExceptionStatus.ProtocolError)
            //    {
            //        Console.WriteLine("Доступ закрыт");
            //    }
            //    else
            //    {
            //        Console.WriteLine($"Ошибка: {ex.ToString()}");
            //    }
            //}
            Console.ReadLine();
         }
    }
}
