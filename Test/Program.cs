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
                    var result = firstPartner.GetUpdataDb().GetAwaiter().GetResult();
                    Console.WriteLine("База обновленна!");
                    Console.WriteLine($"Колличество добавленных записей: {result.Count}");
                }
                else
                    Console.WriteLine("Ничего не скачали :(");
            Console.ReadLine();
        }
    }
}
