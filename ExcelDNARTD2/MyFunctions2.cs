using ExcelDna.Integration;
using System.Text.Json;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using RTD.Excel.Model;

namespace RTD.Excel
{
    public class MyFunctions2
    {
        // 1. Create an instance on the AsyncBatchUtil, passing in some parameters and the btah running function.
        static readonly AsyncBatchUtil BatchRunner = new AsyncBatchUtil(1000, TimeSpan.FromMilliseconds(250), RunBatch);

        // This function will be called for each batch, on a ThreadPool thread.
        // Each AsyncCall contains the function name and arguments passed from the function.
        // The List<object> returned by the Task must contain the results, corresponding to the calls list.
        static async Task<List<object>> RunBatch(List<AsyncBatchUtil.AsyncCall> calls)
        {
            var batchStart = DateTime.Now;
            // Simulate things taking a while...
            //await Task.Delay(TimeSpan.FromSeconds(10));

            // Now build up the list of results...
            var results = new List<object>();

            using (var handler = new HttpClientHandler())
            using (var httpClient = new HttpClient(handler))
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                foreach (var call in calls)
                {
                    string requestUri = (call.Arguments.Length > 0 ? string.Format("http://deckofcardsapi.com/api/deck/{0}/shuffle", call.Arguments[0]) : "http://deckofcardsapi.com/api/deck/new/");
                    HttpResponseMessage getAsync = await httpClient.GetAsync(new Uri(requestUri, UriKind.Absolute));
                    string responseString = await getAsync.Content.ReadAsStringAsync();
                    // As an example just an informative string
                    //var result = string.Format("{0} - {1} : {2}/{3} @ {4:HH:mm:ss.fff}", call.FunctionName, call.Arguments[0], i++, calls.Count, batchStart);
                    results.Add(responseString);
                }                              
            }

            return results;
        }


        [ExcelFunction(Name = "TesteSlowFunction")]
        public static object SlowFunction([ExcelArgument(Name = "RTD")] object RTD)
        {
            var obj = BatchRunner.Run("SlowFunction");

            if (obj.Equals(ExcelError.ExcelErrorNA))
                return "!!! Fetching data";

            return obj;
        }

        [ExcelFunction(Name = "TesteSlowFunctionEmbaralhar")]
        public static object SlowFunctionEmbaralhar([ExcelArgument(Name = "value")] string value)
        {
            Baralho baralho = JsonSerializer.Deserialize<Baralho>(value);

            var obj = BatchRunner.Run("SlowFunctionEmbaralhar", baralho.deck_id);

            if (obj.Equals(ExcelError.ExcelErrorNA))
                return "!!! Fetching data";

            return obj;
        }
    }
}
