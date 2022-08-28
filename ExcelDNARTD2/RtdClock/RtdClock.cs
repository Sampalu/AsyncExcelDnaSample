using ExcelDna.Integration;
using Microsoft.AspNetCore.SignalR.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive.Linq;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;
using System.Threading;
using Newtonsoft.Json;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using RTD.Excel.Model;
using System.Collections;

namespace RTD.Excel
{
    public static class RtdClock
    {
        [ExcelFunction(Description = "Provides a ticking clock")]
        public static object dnaRtdClock_Rx()
        {
            string functionName = "dnaRtdClock";
            object paramInfo = null; // could be one parameter passed in directly, or an object array of all the parameters: new object[] {param1, param2}
            return ObservableRtdUtil.Observe(functionName, paramInfo, () => GetObservableClock());
        }

        static IObservable<string> GetObservableClock()
        {
            return Observable.Timer(dueTime: TimeSpan.Zero, period: TimeSpan.FromSeconds(1))
                             .Select(_ => DateTime.Now.ToString("HH:mm:ss"));
        }

        [ExcelFunction(Description = "Provides a ticking clock")]
        public static object teste_Rx()
        {
            return ArrayResizer2.Resize(GetObservableTeste().Result);
        }

        static async Task<object[,]> GetObservableTeste()
        {
            HubConnection connection = new HubConnectionBuilder()
                .WithUrl("https://localhost:44328/uploadhub")
                .WithAutomaticReconnect()
                .ConfigureLogging(logging =>
                {
                    // Log to the Console
                    logging.AddDebug();

                    logging.SetMinimumLevel(LogLevel.Information);
                    logging.AddFilter("Microsoft.AspNetCore.SignalR", LogLevel.Debug);
                    logging.AddFilter("Microsoft.AspNetCore.Http.Connections", LogLevel.Debug);
                })
                .Build();

            await connection.StartAsync();

            var cancellationTokenSource = new CancellationTokenSource();
            var channel = await connection.StreamAsChannelAsync<object>("DownloadFileAsByteArray", "filePath", false, cancellationTokenSource.Token);

            var notificacao = string.Empty;
            // Wait asynchronously for data to become available
            while (await channel.WaitToReadAsync())
            {
                // Read all currently available data synchronously, before waiting for more data
                while (channel.TryRead(out var notificacaoParte))
                {
                    notificacao += DecodeBase64(System.Text.Encoding.ASCII, notificacaoParte.ToString());
                }
            }


            //connection.On<string>("addFullMessage", (notificacao) =>
            //{
            //    RTD.Excel.Model.AssetNotificacao assetNotificacao =
            //        JsonConvert.DeserializeObject<Dictionary<string, RTD.Excel.Model.AssetNotificacao>>(notificacao).ElementAt(0).Value;


            //});

            RTD.Excel.Model.AssetNotificacao assetNotificacao =
                    JsonConvert.DeserializeObject<Dictionary<string, RTD.Excel.Model.AssetNotificacao>>(notificacao).ElementAt(0).Value;

            return ConverterEmArray(assetNotificacao.Ativo.Payload);
        }

        private static object[,] ConverterEmArray(object asset)
        {
            Dictionary<string, ArrayList> valores = new Dictionary<string, ArrayList>();
            ArrayList colunas = new ArrayList();

            string jsonText = JsonConvert.SerializeObject(asset);
            using (var reader = new JsonTextReader(new StringReader(jsonText)))
            {
                while (reader.Read())
                {
                    if (reader.TokenType == JsonToken.PropertyName)
                    {
                        if (!colunas.Contains(reader.Value))
                        {
                            colunas.Add(reader.Value);
                            //valores
                            valores.Add(reader.Value.ToString(), new ArrayList() { reader.Value.ToString() });
                        }
                    } else if (reader.TokenType == JsonToken.String)
                    {
                        string campo = reader.Path.Split('.')[1];

                        var valor = valores.FirstOrDefault(x => x.Key == campo);
                        valores.Remove(campo);
                        valor.Value.Add(reader.Value);
                        valores.Add(campo, valor.Value);
                    }
                }
            }

            var primeiroItem = valores.FirstOrDefault();
            
            int rows = primeiroItem.Value.Count;
            int columns = valores.Count;

            object[,] result = new object[rows, columns];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    if (i == 0)
                        result[i, j] = colunas[j];
                    else
                    {
                        var valor = valores.FirstOrDefault(x => x.Key == colunas[j].ToString());
                        result[i, j] = valor.Value[i];
                    }
                }
            }

            return result;
        }

        public static string DecodeBase64(this System.Text.Encoding encoding, string encodedText)
        {
            if (encodedText == null)
            {
                return null;
            }

            byte[] textAsBytes = System.Convert.FromBase64String(encodedText);
            return encoding.GetString(textAsBytes);
        }

        public static Object ByteArrayToObject(byte[] arrBytes)
        {
            using (var memStream = new MemoryStream())
            {
                var binForm = new BinaryFormatter();
                memStream.Write(arrBytes, 0, arrBytes.Length);
                memStream.Seek(0, SeekOrigin.Begin);
                var obj = binForm.Deserialize(memStream);
                return obj;
            }
        }
    }
}
