// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.taiwancalendar?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using DotNetLib.Extensions;
using System.Globalization;
using System.Threading.Tasks;

namespace DotNetLib.System.Globalization
{
    public class TaiwanCalendarSingleton : ITaiwanCalendarSingleton
    {
        public TaiwanCalendarSingleton() { }

        // Factory Methods
        public TaiwanCalendar Create()
        {
            return new TaiwanCalendar();
        }

        // Fields
        public int CurrentEra => Calendar.CurrentEra;


        // Methods
        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.TaiwanCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

        public ICalendar ReadOnly(ICalendar pCalendar)
        {
            return Calendar.ReadOnly((Calendar)pCalendar.Unwrap()).Wrap();
        }

    }
}
