using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace DocRandomizer
{
    class Program
    {
#if NETCORE || NET5
    internal const string DocumentDirectory = @"..\..\..\..\Document\";
#else
        internal const string DocumentDirectory = @"..\..\..\Document\";
#endif


        static void Main(string[] args)
        {
            RandTable.Do();
            Console.WriteLine("Готово!");
        }


    }
}