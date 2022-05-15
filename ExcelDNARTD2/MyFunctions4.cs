using System;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelDna.Integration;
using RTD.Excel.Model;

namespace RTD.Excel
{
    public class MyFunctions4
    {
        [ExcelFunction(Name = "TesteSleepPerCaller")]
        public static object TesteSleepPerCaller([ExcelArgument(Name = "seconds")] int seconds)
        {
            var obj = LimitedConcurrencyAsync.AsyncFunctions.SleepPerCaller(seconds);

            // Check the asyncResult to see if we're still busy
            if (obj.Equals(ExcelError.ExcelErrorNA))
                return "!!! Fetching data";

            return obj;
        }

        [ExcelFunction(Name = "TesteSpleep")]
        public static object TesteSpleep([ExcelArgument(Name = "seconds")] int seconds)
        {
            var obj = LimitedConcurrencyAsync.AsyncFunctions.Sleep(seconds);

            // Check the asyncResult to see if we're still busy
            if (obj.Equals(ExcelError.ExcelErrorNA))
                return "!!! Fetching data";

            return obj;
        }

        [ExcelFunction(Name = "TesteLimitedAsync")]
        public static object TesteLimitedAsync([ExcelArgument(Name = "RTD")] string rtd)
        {
            object asyncResult = LimitedConcurrencyAsync.AsyncFunctions.RunBatch1(rtd);

            // Check the asyncResult to see if we're still busy
            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
                return "!!! Fetching data";

            return asyncResult;
        }

        [ExcelFunction(Name = "TesteLimitedAsync2")]
        public static object TesteLimitedAsync2([ExcelArgument(Name = "Deck_id")] string deck_id)
        {
            if (deck_id != "!!! Fetching data")
            {
                object asyncResult = ExcelAsyncUtil.Run("TesteAsyncSemCache2", deck_id, () =>
                {
                    return RunBatch2(deck_id).Result;
                });

                // Check the asyncResult to see if we're still busy
                if (asyncResult.Equals(ExcelError.ExcelErrorNA))
                    return "!!! Fetching data";

                return asyncResult;
            }
            else
            {
                return "!!! Fetching data";
            }
        }

        

        static async Task<object> RunBatch2(string deck_id)
        {
            string responseString = string.Empty;

            using (var handler = new HttpClientHandler())
            using (var httpClient = new HttpClient(handler))
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                string requestUri = string.Format("http://deckofcardsapi.com/api/deck/{0}/shuffle", deck_id);
                HttpResponseMessage getAsync = await httpClient.GetAsync(new Uri(requestUri, UriKind.Absolute));
                responseString = await getAsync.Content.ReadAsStringAsync();
            }

            Baralho baralho = JsonSerializer.Deserialize<Baralho>(responseString);

            return baralho.shuffled;
        }
    }
}
