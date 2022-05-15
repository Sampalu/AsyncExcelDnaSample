using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTD.Excel
{
    public class MyAddIn : IExcelAddIn
    {
        public void AutoClose()
        {          
        }

        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! EXCEPTION: " + ex.ToString());
        }
    }
}
