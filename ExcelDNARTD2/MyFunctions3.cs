using System;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelDna.Integration;
using RTD.Excel.Model;

namespace RTD.Excel
{
    public class MyFunctions3
    {
        [ExcelFunction(Name = "TesteAsyncSemCache")]
        public static object TesteAsyncSemCache([ExcelArgument(Name = "RTD")] string rtd)
        {
            object asyncResult = ExcelAsyncUtil.Run("TesteAsyncSemCache", rtd, () =>
            {
                return RunBatch1().Result;
            });

            // Check the asyncResult to see if we're still busy
            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
                return "!!! Fetching data";

            return asyncResult;
        }

        [ExcelFunction(Name = "TesteAsyncSemCache2")]
        public static object TesteAsyncSemCache2([ExcelArgument(Name = "Deck_id")] string deck_id)
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

        static async Task<object> RunBatch1()
        {
            string responseString = string.Empty;

            using (var handler = new HttpClientHandler())
            using (var httpClient = new HttpClient(handler))
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                string requestUri = "http://deckofcardsapi.com/api/deck/new/";
                HttpResponseMessage getAsync = await httpClient.GetAsync(new Uri(requestUri, UriKind.Absolute));
                responseString = await getAsync.Content.ReadAsStringAsync();
            }

            Baralho baralho = JsonSerializer.Deserialize<Baralho>(responseString);

            return baralho.deck_id;
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


//https://docs.excel-dna.net/asynchronous-functions-with-tasks/
//https://www.codeproject.com/Articles/662009/Streaming-realtime-data-to-Excel
//https://excel-dna.net/2013/03/26/async-and-event-streaming-excel-udfs-with-f/