using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Data.Linq;
using System.Data.Linq.Mapping;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization.Json;
using System.Reflection;
using System.Text;

namespace Test2
{
    public class WorkersList
    {
        public List<Worker> Workers { get; set; }
    }

    [DataContract]
    public class Worker
    {
        [DataMember]
        string str;
        [DataMember]
        public int Id { get; set; }
        [DataMember]
        public int Degrees { get; set; }
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public string Email { get; set; }
        [DataMember]
        public int Selery { get; set; }
        public Worker(string _str)
        {
            str = _str;
        }
    }

    public class DegreesList
    {
        public List<Degrees> Degrees { get; set; }
    }

    public class Degrees
    {
        public int Id { get; set; }
        public string Degree { get; set; }
    }

    public class MyClass
    {

    }

    public class MyClass2: MyClass
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string SecondName { get; set; }
        public MyClass mc;
        public void Print ()
        {

        }
    }

    interface IPoint
    {

    }

    class Program
    {     

        public class MyIExeption : ApplicationException
        {
            public int WhatIsI { get; set; }

            public MyIExeption(string message, int i) : base(message)
            {
                this.WhatIsI = i;
            }
        }

        public T Cast<T>(object input, string str = "cool")
        {
            return (T)input;
        }

        static void Sum(params int[] param)
        {
            int i = 0;
            foreach (var item in param)
            {
                i += item;
            }
            Console.WriteLine("Sum = {0}",i);
        }

        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("------------------Begin-----------------");
                List<string> strList = new List<string>();
                //strList.Add("first");
                //strList.Add("second");
                //strList.Add("third");
                //strList.Add("forth");
                foreach (var item in strList)
                {
                    Console.WriteLine(item);
                }
                Console.WriteLine("******************************");
                foreach (var item in strList.AsEnumerable().Reverse())
                {
                    Console.WriteLine(item);
                }

            }
            catch (MyIExeption ex)
            {
                Console.WriteLine($@"Error:
    Message {ex.Message}
    i is: {ex.WhatIsI}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message} Error From: {ex.StackTrace}");
                Console.WriteLine($"Target:  {ex.TargetSite}");
                Console.WriteLine(ex is SystemException);
                if (ex.InnerException != null)
                    Console.WriteLine($"Inner Exeption: {ex.InnerException.GetBaseException().Message}");
            }
            Console.WriteLine("------------------Done------------------");
            Console.ReadLine();
        }
    }
}
