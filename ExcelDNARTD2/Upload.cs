using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Channels;
using System.Threading.Tasks;
using System.Threading;
using Newtonsoft.Json;
using RTD.Excel.Model;

namespace RTD.Excel
{
    public class Upload
    {
        public ChannelReader<object> UploadAssetAsByteArray(string filePath, bool throwException, CancellationToken cancellationToken)
        {
            var json = "";// JsonFilesRepo.Files["ativo.json"];

            var asset = JsonConvert.DeserializeObject<Asset>(json);

            AssetNotificacao notificacao = new AssetNotificacao(asset);

            Dictionary<string, AssetNotificacao> dicionario = new Dictionary<string, AssetNotificacao>
            {
                { "CHAVE", notificacao }
            };

            var channel = Channel.CreateUnbounded<object>();
            _ = WriteToChannelAsByteArray(channel.Writer, dicionario, throwException, cancellationToken);
            return channel.Reader;
        }

        private async Task WriteToChannelAsByteArray(ChannelWriter<object> writer, Dictionary<string, AssetNotificacao> dicionario, bool throwException, CancellationToken cancellationToken)
        {
            Exception localException = null;
            int countOfChunks = 0;
            int chunkSizeInKb = 60;
            try
            {
                string output = JsonConvert.SerializeObject(dicionario);

                int chunkSize = (int)(chunkSizeInKb * 10);
                int chunkLength = output.Length > chunkSize ? chunkSize : output.Length;

                byte[] array = Encoding.ASCII.GetBytes(output);

                int numOfChunks = Convert.ToInt32(Math.Ceiling(Convert.ToDecimal(array.Length) / Convert.ToDecimal(chunkLength)));

                for (int i = 0; i < array.Length; i += chunkLength)
                {
                    if (countOfChunks == numOfChunks - 1)
                    {
                        chunkLength = output.Length - i;
                    }

                    byte[] b = new byte[chunkLength];
                    Buffer.BlockCopy(array, i, b, 0, b.Length);

                    /* this delay has been added to throttle the rate at which items are written to the channel. 
                    If not throttled, all chunks are written to the channel rapidly before cancellationToken is received */
                    //await Task.Delay(10);
                    if (cancellationToken.IsCancellationRequested)
                    {
                        //Logger.LogInformation("Upstream: streaming cancelled");
                        return;
                    }
                    //Logger.LogInformation($"chunk sent: {Convert.ToBase64String(b)}");
                    await writer.WriteAsync(b, cancellationToken);
                    countOfChunks++;
                }
                //Logger.LogInformation($"Upstream: Total {numOfChunks} chunks written to channel");
            }
            catch (Exception ex)
            {
                localException = ex;
                //Logger.LogError("Error occurred while writing to channel, Exception: " + ex.Message);
            }
            finally
            {
                if (throwException && localException == null)
                {
                    /* Due to this issue: https://github.com/dotnet/aspnetcore/issues/33753, currently only OperationCanceledException
                     are caught by SignalR.*/
                    localException = new OperationCanceledException("Custom Exception thrown from service");
                }
                writer.TryComplete(localException);
            }
        }
    }
}
