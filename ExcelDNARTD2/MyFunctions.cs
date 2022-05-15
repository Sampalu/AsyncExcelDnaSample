using System;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Reactive.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using RTD.Excel;
using RTD.Excel.Model;

namespace RTD.Excel1
{
    public static class MyFunctions
    {
        [ExcelFunction(Name = "SayHello")]
        public static string SayHello(string name)
        {
            return "Hello " + name + "!";
        }

        [ExcelFunction(Name = "TesteCachedAsync")]
        public static object TesteCachedAsync([ExcelArgument(Name = "RTD")] string rtd)
        {
            return RTD.Excel.dnaFunctions.dnaCachedAsync(rtd);
        }        
    }
}
