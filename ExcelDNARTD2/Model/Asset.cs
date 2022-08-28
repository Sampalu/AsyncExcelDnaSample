using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTD.Excel.Model
{
    public class Asset
    {
        public Asset()
        {

        }

        [JsonProperty("endpoint")]
        public string Endpoint { get; set; }

        [JsonProperty("publication_id")]
        public string PublicationID { get; set; }

        [JsonProperty("payload")]
        public object Payload { get; set; }
    }
}
