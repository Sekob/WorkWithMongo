using System;
using System.Net;
using Ionic.Zip;
using System.IO;
using System.Xml.Serialization;
using System.Text;
using System.Xml;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Threading;
using MongoDB.Driver;
using System.Configuration;
using MongoDB.Bson;

namespace UpdateDb
{
    public class DataUpdater
    {
        string _loginName;
        string _loginPwd;
        string _partnerId;
        public Worksheet Data { get; set; }
        public string LoginName { get { return _loginName; } set { _loginName = value; } }
        public string LoginPwd{ get { return _loginPwd; } set { _loginPwd = value; } }
        public string PartnerId { get; set; }

        public class DataContext
        {
            public const string CONNECTION_STRING_NAME = "mongodb://localhost:27017";
            public const string DATABASE_NAME = "Boo";
            public const string WORK_COLLECTION_NAME = "eo_clients";

            private static readonly IMongoClient _client;
            private static readonly IMongoDatabase _database;

            static DataContext()
            {
                _client = new MongoClient(CONNECTION_STRING_NAME);
                _database = _client.GetDatabase(DATABASE_NAME);
            }

            public IMongoClient Client
            {
                get { return _client; }
            }

            public IMongoCollection<EOData> EoClients
            {
                get { return _database.GetCollection<EOData>(WORK_COLLECTION_NAME); }
            }
        }

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

        public static async Task<CookieContainer> GetCookiesAsync(string url, string name, string password)
        {
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.CookieContainer = new CookieContainer();
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            try
            {
                using (var requestStream = await request.GetRequestStreamAsync())
                using (var writer = new StreamWriter(requestStream))
                {
                    writer.Write($"Name={name}&Password={password}&returnUrl=");
                }
                using (var responseStream = await request.GetResponseAsync())
                    return request.CookieContainer;
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError)
                {
                    Console.WriteLine("Доступ закрыт");
                }
                else
                {
                    Console.WriteLine(ex.Message);
                }
                return null;
            }
        }

        public DataUpdater(string loginName, string loginPwd, string partnerId)
        {
            _loginName = loginName;
            _loginPwd = loginPwd;
            _partnerId = partnerId;
        }

        public DataUpdater()
        {
            _loginName = "teleskopsoft";
            _loginPwd = "118";
            _partnerId = "118";
        }

        //TODO разобраться с возвратом
        public async Task<bool> GetDataFromServer()
        {
            ////Working part that gets cookie and html documetn from web
            //var cookie = await GetCookiesAsync("http://bo-otchet.1gl.ru/Account/Login", _loginName, _loginPwd);
            

            ////Getting data from web by WebClient
            //using (var client = new CookieAwareWebClient(cookie))
            //{
            //    Uri newUri = new Uri("http://bo-otchet.1gl.ru/Registration/RepostRegistrationExcel?Caption=%D0%A1%20%D0%B2%D1%8B%D0%BF%D1%83%D1%89%D0%B5%D0%BD%D0%BD%D0%BE%D0%B9%20%D0%9A%D0%AD%D0%9F&ShowSearchForm=False&PageId=0&PageSize=50&SuppressFlags=0&ViewId=HaveCertificate&PartnerId="+PartnerId);
            //    client.DownloadFile (newUri, $"{_partnerId}.xlsx");
            //    using (ZipFile zip = ZipFile.Read($"{_partnerId}.xlsx"))
            //    {
            //        if (zip == null) Console.WriteLine("Null");
            //        else
            //        {
            //            foreach (var zp in zip)
            //            {
            //                zp.Extract(ExtractExistingFileAction.OverwriteSilently);
            //            }
            //        }
            //    }
            //}
            try
            {
                XmlDocument _xml = new XmlDocument();
                try
                {
                    _xml.Load("xl\\worksheets\\sheet1.xml");
                }
                catch (System.IO.FileNotFoundException ex)
                {
                    Console.WriteLine($"File not found!{Environment.NewLine} Error: {ex.ToString()}");
                }
                string str = _xml.OuterXml;
                str = str.Replace("<v>", "").Replace("</v>", "")
                         .Replace("<f>=TEXT", "").Replace("</f>", "");

                XmlRootAttribute xRoot = new XmlRootAttribute();
                xRoot.ElementName = "worksheet";
                xRoot.Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                xRoot.IsNullable = true;
                var xml = new XmlSerializer(typeof(Worksheet),xRoot);
                Data = xml.Deserialize(new MemoryStream(Encoding.UTF8.GetBytes(str))) as Worksheet;
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
         }
        //TODO обработать ошибки
        public async Task<List<EOData>> AddEoClientsDb()
        {
            List < EOData > dataResult = new List<EOData>();
            var dataContext = new DataContext();
            var dataFromDB = await dataContext.EoClients.Find(new BsonDocument()).ToListAsync();
            if (dataFromDB.Count!=0)
            {
                foreach (var item in Data.Rows)
                {
                    var result = dataContext.EoClients.Find(x => x._id == item._id).FirstOrDefault();
                    if (result==null)
                    {
                        dataContext.EoClients.InsertOne(item);
                        dataResult.Add(item);
                    }
                }
            }
            else
            {
                foreach (var item in Data.Rows)
                {
                    dataContext.EoClients.InsertOne(item);
                    dataResult.Add(item);
                }
            }
            return dataResult;
        }
        //TODO обработать ошибки
        public async Task<List<EOData>> CheckAndUpdateForClients()
        {
            List<EOData> updatedData = new List<EOData>();
            object _lock = new object();
            var dataContext = new DataContext();
            var result = await dataContext.EoClients.Find(x => true).ToListAsync();
            Parallel.ForEach(Data.Rows, item => {
                foreach (var itm in result)
                {
                    if (itm._id==item._id)
                    {
                        if (!itm.Equals(item))
                        {
                            lock (_lock)
                            {
                                dataContext.EoClients.ReplaceOne<EOData>((x) => x._id == itm._id, item);
                                updatedData.Add(item);
                            }
                        }
                    }
                }
            });
            return updatedData;
        }
        //TODO обработать ошибки
        public async Task<List<EOData>> CheckUpdateForClients()
        {
            List<EOData> updatedData = new List<EOData>();
            object _lock = new object();
            var dataContext = new DataContext();
            var result = await dataContext.EoClients.Find(x => true).ToListAsync();
            Parallel.ForEach(Data.Rows, item => {
                foreach (var itm in result)
                {
                    if (itm._id == item._id)
                    {
                        if (!itm.Equals(item))
                            lock (_lock)
                                updatedData.Add(item);
                    }
                }
            });
            return updatedData;
        }

        static void Main (string[] args)
        {

        }
    }
}
