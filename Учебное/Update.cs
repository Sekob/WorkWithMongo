using System;
using System.Net;
using System.IO;
using System.Xml.Serialization;
using System.Text;
using System.Xml;
using System.Collections.Generic;
using System.Threading.Tasks;
using MongoDB.Driver;
using MongoDB.Bson;
using Ionic.Zip;
using System.Threading;

namespace UpdateDb
{
    
    public class DataContext
    {
        public const string CONNECTION_STRING_NAME = "mongodb://localhost:27017";
        public const string DATABASE_NAME = "Boo";
        public const string WORK_COLLECTION_NAME = "eo_clients";
        public const string TASK_COLLECTION_NAME = "tasks";

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
        public IMongoCollection<DbWorkingTasks> WorkingTasks
        {
            get { return _database.GetCollection<DbWorkingTasks>(TASK_COLLECTION_NAME); }
        }
        public IMongoCollection<EOData> EoClients
        {
            get { return _database.GetCollection<EOData>(WORK_COLLECTION_NAME); }
        }
    }

    public class DataUpdater
    {
        string _loginName;
        string _loginPwd;
        string _partnerId;
        public Worksheet Data { get; set; }
        public string LoginName { get { return _loginName; } set { _loginName = value; } }
        public string LoginPwd { get { return _loginPwd; } set { _loginPwd = value; } }
        public string PartnerId { get; set; }

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
            ShowMessageClass showMessage = new ShowConsoleMessage();
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
                    showMessage.ShowMessage("Доступ закрыт");
                }
                else
                {
                    showMessage.ShowMessage(ex.Message);
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

        //TODO: разобраться с возвратом
        //Подключается к серверу, скачевает xlsx фаил, разархивируетв в xml и десереализует/
        public async Task<bool> GetDataFromServer()
        {
            ////Working part that gets cookie and html documetn from web
            var cookie = await GetCookiesAsync("http://bo-otchet.1gl.ru/Account/Login", _loginName, _loginPwd);


            //Getting data from web by WebClient
            using (var client = new CookieAwareWebClient(cookie))
            {
                Uri newUri = new Uri("http://bo-otchet.1gl.ru/Registration/RepostRegistrationExcel?Caption=%D0%A1%20%D0%B2%D1%8B%D0%BF%D1%83%D1%89%D0%B5%D0%BD%D0%BD%D0%BE%D0%B9%20%D0%9A%D0%AD%D0%9F&ShowSearchForm=False&PageId=0&PageSize=50&SuppressFlags=0&ViewId=HaveCertificate&PartnerId=" + PartnerId);
                client.DownloadFile(newUri, $"{_partnerId}.xlsx");
                using (ZipFile zip = ZipFile.Read($"{_partnerId}.xlsx"))
                {
                    if (zip == null) Console.WriteLine("Null");
                    else
                    {
                        foreach (var zp in zip)
                        {
                            zp.Extract(ExtractExistingFileAction.OverwriteSilently);
                        }
                    }
                }
            }
            ShowMessageClass showMessage = new ShowConsoleMessage();
            try
            {
                XmlDocument _xml = new XmlDocument();
                try
                {
                    _xml.Load("xl\\worksheets\\sheet1.xml");
                }
                catch (System.IO.FileNotFoundException ex)
                {
                    showMessage.ShowMessage($"File not found!{Environment.NewLine} Error: {ex.ToString()}");
                }
                string str = _xml.OuterXml;
                str = str.Replace("<v>", "").Replace("</v>", "")
                         .Replace("<f>=TEXT", "").Replace("</f>", "");

                XmlRootAttribute xRoot = new XmlRootAttribute();
                xRoot.ElementName = "worksheet";
                xRoot.Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                xRoot.IsNullable = true;
                var xml = new XmlSerializer(typeof(Worksheet), xRoot);
                Data = xml.Deserialize(new MemoryStream(Encoding.UTF8.GetBytes(str))) as Worksheet;
                return true;
            }
            catch (Exception ex)
            {
                showMessage.ShowMessage(ex.Message);
                return false;
            }
        }

        //TODO: обработать ошибки
        public async Task<List<EOData>> AddEoClientsDb()
        {
            if (Data == null) throw new Exception($"Ошибка: GetDataFromServer(): Данные с сервера не были получены!\nИнициализируйте GetDataFromServer()");
            List<EOData> dataResult = new List<EOData>();
            var dataContext = new DataContext();
            var dataFromDB = await dataContext.EoClients.Find(new BsonDocument()).ToListAsync();
            if (dataFromDB.Count != 0)
            {
                foreach (var item in Data.Rows)
                {
                    var result = dataContext.EoClients.Find(x => x._id == item._id).FirstOrDefault();
                    if (result == null)
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
        //TODO: обработать ошибки
        public async Task<List<EOData>> CheckAndUpdateForClients()
        {
            if (Data == null) throw new Exception($"Ошибка: CheckAndUpdateForClients(): Данные с сервера не были получены!\nИнициализируйте GetDataFromServer()");
            List<EOData> updatedData = new List<EOData>();
            object _lock = new object();
            var dataContext = new DataContext();
            var result = await dataContext.EoClients.Find(x => true).ToListAsync();
            Parallel.ForEach(Data.Rows, item =>
            {
                foreach (var itm in result)
                {
                    if (itm._id == item._id)
                    {
                        if (!itm.Equals(item))
                        {
                            lock (_lock)
                            {
                                var update = Builders<EOData>.Update.Set((x) => x.Accountent, item.Accountent)
                                                                    .Set((x) => x.Agent, item.Agent)
                                                                    .Set((x) => x.Name, item.Name)
                                                                    .Set((x) => x.INN, item.INN)
                                                                    .Set((x) => x.Email, item.Email)
                                                                    .Set((x) => x.Mobile, item.Mobile)
                                                                    .Set((x) => x.ES, item.ES)
                                                                    .Set((x) => x.PFR, item.PFR)
                                                                    .Set((x) => x.Notification, item.Notification)
                                                                    .Set((x) => x.Partner, item.Partner)
                                                                    .Set((x) => x.LastUpdate, item.LastUpdate)
                                                                    .Set((x) => x.DateStart, item.DateStart)
                                                                    .Set((x) => x.DateEnd, item.DateEnd)
                                                                    .Set((x) => x.Type, item.Type);
                                                                    
                                dataContext.EoClients.UpdateOne<EOData>(x => x._id == itm._id, update);
                                updatedData.Add(item);
                            }
                        }
                        break;
                    }
                }
            });
            return updatedData;
        }
        //TODO: обработать ошибки
        public async Task<List<EOData>> CheckUpdateForClients()
        {
            if (Data == null) throw new Exception($"Ошибка: CheckUpdateForClients(): Данные с сервера не были получены!\nИнициализируйте GetDataFromServer()");
            List<EOData> updatedData = new List<EOData>();
            object _lock = new object();
            var dataContext = new DataContext();
            var result = await dataContext.EoClients.Find(x => true).ToListAsync();
            Parallel.ForEach(Data.Rows, item =>
            {
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
    }


    public class MyReplaceOneResult 
    {
        long _matchesCount;
        long _modifiedCount;
        public long MatchedCount
        {
            get
            {
                return _matchesCount;
            }
            set
            {
                _matchesCount = value;
            }
        }
        public long ModifiedCount
        {
            get
            {
                return _modifiedCount;
            }
            set
            {
                _modifiedCount = value;
            }
        }
    }

    public class StaticDBMethods
    {
        //TODO: Добавить проверку на не корректные Value
        public static async Task<MyReplaceOneResult> DoUpdateEoClients(List<EOData> _data)
        {
            MyReplaceOneResult result = new MyReplaceOneResult();
            var dataContext = new DataContext();
            try
            {
                foreach (var item in _data)
                {
                    var update = Builders<EOData>.Update.Set((x) => x.Accountent, item.Accountent)
                                                                    .Set((x) => x.Agent, item.Agent)
                                                                    .Set((x) => x.Name, item.Name)
                                                                    .Set((x) => x.INN, item.INN)
                                                                    .Set((x) => x.Email, item.Email)
                                                                    .Set((x) => x.Mobile, item.Mobile)
                                                                    .Set((x) => x.ES, item.ES)
                                                                    .Set((x) => x.PFR, item.PFR)
                                                                    .Set((x) => x.Notification, item.Notification)
                                                                    .Set((x) => x.Partner, item.Partner)
                                                                    .Set((x) => x.LastUpdate, item.LastUpdate)
                                                                    .Set((x) => x.DateStart, item.DateStart)
                                                                    .Set((x) => x.DateEnd, item.DateEnd)
                                                                    .Set((x) => x.Type, item.Type);

                    var temp = await dataContext.EoClients.UpdateOneAsync<EOData>(x => x._id == item._id, update);
                    result.MatchedCount += temp.MatchedCount;
                    result.ModifiedCount += temp.ModifiedCount;
                }
            }
            catch (Exception)
            {
                throw;
            }
            return result;
        }
        public static async Task<bool> AddMessageToEoClient(string shortID, string message)
        {
            ShowMessageClass showMessage = new ShowConsoleMessage();
            var dataContex = new DataContext();
            var _message = new EoMessages();
            _message.From = "Kobelev"; // от кого
            _message.To = "Handogka"; // кому
            _message.Message = message;
            _message.PutDate = DateTime.Now.ToLocalTime();
            var findResult = await dataContex.EoClients.Find<EOData>((x) => x.Short_Id == shortID).FirstOrDefaultAsync();
            if (findResult.Message == null)
            {
                findResult.Message = new List<EoMessages>();
                findResult.Message.Add(_message);
                var replaceResult = await dataContex.EoClients.ReplaceOneAsync<EOData>((x) => x.Short_Id == findResult.Short_Id, findResult);
                if (replaceResult.ModifiedCount != 0)
                {
                    showMessage.ShowMessage("Сообщеине добавленно");
                    return true;
                }
                else
                {
                    showMessage.ShowMessage("Сообщеине добавленно");
                    return false;
                }
            }
            else
            {
                var builder = Builders<EOData>.Update.Push("Message", _message);
                var result = await dataContex.EoClients.UpdateOneAsync<EOData>((x) => x.Short_Id == shortID, builder);
                if (result.ModifiedCount != 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
        public static async Task<bool> NewEoTask(string _title, string _task, string _from, string _to, string _shortId, double _hours)
        {
            ShowMessageClass message = new ShowConsoleMessage();
            DataContext dataContex = new DataContext();
            DbWorkingTasks newTask = new DbWorkingTasks
            {
                _id = ObjectId.GenerateNewId(),
                ShortEoClientId = _shortId,
                Title = _title,
                PutTime = DateTime.Now.ToLocalTime(),
                TillDone = DateTime.Now.AddHours(_hours),
                To = _to,
                From = _from,
                IsDone = false,
                Messages = new List<EoMessages>(),
                Work = new List<EoWork>(),
                Task = _task
            };
            try
            {
                await dataContex.WorkingTasks.InsertOneAsync(newTask);
                message.ShowMessage($"Задача {_to} для {_shortId} поставлена!");
                return true;
            }
            catch (Exception)
            {
                throw;
            }

        }
        //TODO: Разобраться с отловлей ошибок
        public static async Task<List<DbWorkingTasks>> GetTasks()
        {
            DataContext dataContex = new DataContext();
            ShowMessageClass message = new ShowConsoleMessage();
            try
            {
                var result = await dataContex.WorkingTasks.Find<DbWorkingTasks>((x) => true).ToListAsync();
                if (result.Count == 0)
                    throw new NotImplementedException("Список задач пуст!");
                return result;
            }
            catch (MongoException ex)
            {
                throw new NotImplementedException($"Внимание!MongoDb ОШИБКА {ex.HResult}: {ex.Message}\nДанные: {ex.Data}");
            }
        }
        public static async Task<List<DbWorkingTasks>> GetClientsTasks(string _shortId)
        {
            DataContext dataContex = new DataContext();
            ShowMessageClass message = new ShowConsoleMessage();
            try
            {
                var result = await dataContex.WorkingTasks.Find<DbWorkingTasks>((x) => x.ShortEoClientId==_shortId).ToListAsync();
                if (result.Count == 0)
                    throw new NotImplementedException("Мы не смогли найти клиента с таким ID!");
                return result;
            }
            catch (MongoException ex)
            {
                throw new NotImplementedException($"Внимание!MongoDb ОШИБКА {ex.HResult}: {ex.Message}\nДанные: {ex.Data}");            
            }
        }
    }
}
