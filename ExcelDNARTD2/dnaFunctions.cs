using System;
using System.Threading;
using System.Runtime.Caching;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Net.Http;
using System.Net;
using RTD.Excel.Model;
using System.Text.Json;

namespace RTD.Excel
{
    public static class dnaFunctions
    {
        public static object dnaCachedAsync(string input)
        {
            // First check the cache, and return immediately 
            // if we found something.
            // (We also need a unique key to identify the cache item)
            string key = "dnaCachedAsync:" + input;
            ObjectCache cache = MemoryCache.Default;
            string cachedItem = cache[key] as string;
            if (cachedItem != null)
                return cachedItem;

            // Not in the cache - make the async call 
            // to retrieve the item. (The second parameter here should identify 
            // the function call, so would usually be an array of the input parameters, 
            // but here we have the identifying key already.)
            object asyncResult = ExcelAsyncUtil.Run("dnaCachedAsync", key, () =>
            {
                // Here we fetch the data from far away....
                // This code will run on a ThreadPool thread.

                // To simulate a slow calculation or web service call,
                // Just sleep for a few seconds...
                //Thread.Sleep(5000);               
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

                // Then return the result
                //return "The calculation with input "
                //        + input + " completed at "
                //        + DateTime.Now.ToString("HH:mm:ss");
            });

            // Check the asyncResult to see if we're still busy
            if (asyncResult.Equals(ExcelError.ExcelErrorNA))
                return "!!! Fetching data";

            // OK, we actually got the result this time.
            // Add to the cache and return
            // (keeping the cached entry valid for 1 minute)
            // Note that the function won't recalc automatically after 
            //    the cache expires. For this we need to go the 
            //    RxExcel route with an IObservable.
            cache.Add(key, asyncResult, DateTime.Now.AddMinutes(1), null);
            return asyncResult;
        }

        public static string dnaTest()
        {
            return "Hello from CachedAsyncSample";
        }
    }
}
