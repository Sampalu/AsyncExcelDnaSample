using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Schedulers;
using ExcelDna.Integration;
using ExcelDna.Utils;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using RTD.Excel.Model;
using System.Collections.Generic;

namespace LimitedConcurrencyAsync
{
    public static class AsyncFunctions
    {
        static TaskFactory _fourThreadFactory;

        static AsyncFunctions()
        {
            // This initialization could be lazy (and of course be any other TaskScheduler)
            var fourThreadScheduler = new LimitedConcurrencyLevelTaskScheduler(4);
            _fourThreadFactory = new TaskFactory(fourThreadScheduler);
        }

        public static object Sleep(int seconds)
        {
            Debug.Print($"Sleep Call: {seconds}");
            // The callerFunctionName and callerParameters are internally combined and used as a 'key' 
            // to link the underlying RTD calls together.
            string callerFunctionName = "Sleep";
            object callerParameters = new object[] {seconds}; // This need not be an array if it's just a single parameter

            var result = AsyncTaskUtil.RunAsTask(callerFunctionName, callerParameters, _fourThreadFactory, () =>
                {
                    Thread.Sleep(seconds * 1000);
                    return "Slept on Thread " + Thread.CurrentThread.ManagedThreadId;
                });
            Debug.Print($"Sleep Result: {result}");
            return result;
        }

        public static object SleepPerCaller(int seconds)
        {
            // Trick to get each call to be a separate instance
            // Normally you only want to add the actual parameters passed in
            object callerReference = XlCall.Excel(XlCall.xlfCaller);
            string callerFunctionName = "SleepPerCaller";
            object callerParameters = new object[] { seconds, callerReference };

            // The RunTask version (instead of RunAsTask used above) is more flexible if the Task will be created in some other way.
            return AsyncTaskUtil.RunTask(callerFunctionName, callerParameters, () =>
            {
                // The function here should return the Task to run
                return _fourThreadFactory.StartNew(() =>
                {
                    Thread.Sleep(seconds * 1000);
                    return string.Format("Slept on Thread {0}, called from {1}", Thread.CurrentThread.ManagedThreadId, callerReference);
                });
            });
        }

        public static object RunBatch1(string param)
        {
            string callerFunctionName = "RunBatch1";
            object callerParameters = new object[] { param }; // This need not be an array if it's just a single parameter

            var result = AsyncTaskUtil.RunAsTask(callerFunctionName, callerParameters, _fourThreadFactory, () =>
            {
                string responseString = string.Empty;

                using (var handler = new HttpClientHandler())
                using (var httpClient = new HttpClient(handler))
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                    string requestUri = "http://deckofcardsapi.com/api/deck/new/";
                    HttpResponseMessage getAsync = httpClient.GetAsync(new Uri(requestUri, UriKind.Absolute)).Result;
                    responseString = getAsync.Content.ReadAsStringAsync().Result;
                }

                Baralho baralho = JsonSerializer.Deserialize<Baralho>(responseString);

                return baralho.deck_id;
            });
            
            return result;
        }

        public static object SumPageSizesAsync(string param)
        {
            string callerFunctionName = "SumPageSizesAsync";
            object callerParameters = new object[] { param }; // This need not be an array if it's just a single parameter

            //CancellationTokenSource cts = new CancellationTokenSource();

            var result = AsyncTaskUtil.RunAsTaskWithCancellation(callerFunctionName, callerParameters, _fourThreadFactory, (token) =>
            {
                string responseString = string.Empty;

                HttpClient httpClient = new HttpClient
                {
                    MaxResponseContentBufferSize = 1_000_000
                };

                IEnumerable<string> urlList = new string[]
                {
                    "https://docs.microsoft.com",
                    "https://docs.microsoft.com/aspnet/core",
                    "https://docs.microsoft.com/azure",
                    "https://docs.microsoft.com/azure/devops",
                    "https://docs.microsoft.com/dotnet",
                    "https://docs.microsoft.com/dynamics365",
                    "https://docs.microsoft.com/education",
                    "https://docs.microsoft.com/enterprise-mobility-security",
                    "https://docs.microsoft.com/gaming",
                    "https://docs.microsoft.com/graph",
                    "https://docs.microsoft.com/microsoft-365",
                    "https://docs.microsoft.com/office",
                    "https://docs.microsoft.com/powershell",
                    "https://docs.microsoft.com/sql",
                    "https://docs.microsoft.com/surface",
                    "https://docs.microsoft.com/system-center",
                    "https://docs.microsoft.com/visualstudio",
                    "https://docs.microsoft.com/windows",
                    "https://docs.microsoft.com/xamarin"
                };

                int total = 0;
                foreach (string url in urlList)
                {
                    var contentLengthString = ProcessUrl(url, httpClient, token);
                    if (!Int32.TryParse(contentLengthString.ToString(), out int contentLength))
                        return contentLengthString;
                    total += contentLength;
                }                

                return total;
            });

            return result;
        }

        private static object ProcessUrl(string url, HttpClient client, CancellationToken token)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                HttpResponseMessage response = client.GetAsync(url, token).Result;
                byte[] content = response.Content.ReadAsByteArrayAsync().Result;
                Console.WriteLine($"{url,-60} {content.Length,10:#,#}");

                return content.Length;
            }
            catch (AggregateException)
            {
                return "Tasks cancelled: timed out.";
            }
        }
    }
}
