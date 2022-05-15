using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
