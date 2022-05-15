using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTD.Excel
{
    public class DataWrapper
    {
        public DataWrapper()
        {

        }
        public DataWrapper(string texto)
        {
            Data = texto;
        }
        public object Data { get; set; }
    }
}
