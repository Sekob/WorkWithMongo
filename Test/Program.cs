using System;
using UpdateDb;
using MongoDB.Driver;
using MongoDB.Bson;
using System.Threading;
using System.Collections.Generic;
using System.Diagnostics;

namespace Test
{
    class Program
    {
        static Timer task_timer;
        static List<DbWorkingTasks> _tasks = new List<DbWorkingTasks>();
        static void Main(string[] args)
        {
            ShowMessageClass message = new ShowConsoleMessage();
            try
            {
                task_timer = new Timer(CheckTasks, null, Timeout.Infinite, Timeout.Infinite);
                task_timer.Change(0, Timeout.Infinite);
                Console.ReadLine();
                task_timer.Dispose();
            }
            catch (Exception ex)
            {
                message.ShowMessage(ex.Message);
            }
            Console.ReadLine();
        }

        static void CheckTasks (Object obj)
        {
             
            var list = StaticDBMethods.GetTasks().GetAwaiter().GetResult();
            if (_tasks.Count<list.Count)
            {
                Console.WriteLine("//////////////////////////////////");
                foreach (var item in list)
                {
                    _tasks.Add(item);
                    Console.WriteLine($"From: {item.From} To: {item.To}");
                    Console.WriteLine($"When: {item.PutTime.ToShortTimeString()} Till: {item.TillDone.ToShortTimeString()}");
                    Console.WriteLine($"Title: {item.Title}");
                    Console.WriteLine($"Task: {item.Task}\nIs Done: {item.IsDone.ToString()}");
                }
            }            
            task_timer.Change(10000,Timeout.Infinite);
        }
    }
}
