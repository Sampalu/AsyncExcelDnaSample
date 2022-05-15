using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reactive.Linq;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;
using Microsoft.AspNetCore.SignalR.Client;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace RTD.Excel1
{
    [ComVisible(true)]                   // Required since the default template puts [assembly:ComVisible(false)] in the AssemblyInfo.cs
    [ProgId(SignalRServer.ServerProgId)]     //  If ProgId is not specified, change the XlCall.RTD call in the wrapper to use namespace + type name (the default ProgId)
    public class SignalRServer : ExcelRtdServer
    {
        public const string ServerProgId = "ExcelDNARTD2.SignalRServer";
        private HubConnection _connection;

        public SignalRServer()
        {
            _connection = new HubConnectionBuilder()
                .WithUrl("https://localhost:5001/hub/chat")
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

            _connection.StartAsync();

            _connection.On<string>("addMessage", (endpointDataId) =>
            {
                //var newMessage = $"{message}";
                foreach (Topic topic in _topics)
                {
                    topic.UpdateValue(endpointDataId);
                }
            });
        }

        // Using a System.Threading.Time which invokes the callback on a ThreadPool thread 
        // (normally that would be dangeours for an RTD server, but ExcelRtdServer is thread-safe)
       Timer _timer;
        List<Topic> _topics;

        protected override bool ServerStart()
        {
            _timer = new Timer(timer_tick, null, 0, 1000);
            _topics = new List<Topic>();
            return true;
        }

        protected override void ServerTerminate()
        {
           _timer.Dispose();            
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            _topics.Add(topic);
            return DateTime.Now.ToString("HH:mm:ss");
        }

        protected override void DisconnectData(Topic topic)
        {
            _topics.Remove(topic);
        }

        void timer_tick(object _unused_state_)
        {
            string now = DateTime.Now.ToString("HH:mm:ss");
            foreach (var topic in _topics)
                topic.UpdateValue(now);
        }

        private string RetornarStatusConexao()
        {
            string conexao = "...";
            if (_connection != null)
                conexao = _connection.State.ToString();

            return conexao;
        }
    }
}