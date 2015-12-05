using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UpdateDb;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            DataUpdater firstPartner = new DataUpdater("teleskopsoft", "118", "118");
            if (firstPartner.GetDataFromServer().GetAwaiter().GetResult())
                if (firstPartner.Data != null)
                {
                    var addResult = firstPartner.AddEoClientsDb().GetAwaiter().GetResult();
                    Console.WriteLine("База обновленна!");
                    Console.WriteLine($"Колличество добавленных записей: {addResult.Count}");
                    var updateResult = firstPartner.CheckUpdateForClients().GetAwaiter().GetResult();
                    if (updateResult.Count!=0)
                    {
                        Console.WriteLine($"Количество записей нуждающихся в обновлении: {updateResult.Count}");
                        foreach (var item in updateResult)
                        {
                            Console.WriteLine("Обновленная запись:");
                            Console.WriteLine(item.Print());
                        }
                    }
                    else
                    {
                        Console.WriteLine("Записи не нуждаются в обновлении!");
                    }
                }
                else
                    Console.WriteLine("Ничего не скачали :(");
            Console.ReadLine();
        }
    }
}
