// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.taiwancalendar?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using DotNetLib.Extensions;
using System.Globalization;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("ED1EF286-3970-41C5-BC25-91FD754FACB3")]
    [ProgId("DotNetLib.System.Globalization.PersianCalendarSingleton")]
    [Description("Represents the Taiwan calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITaiwanCalendarSingleton))]
    public class TaiwanCalendarSingleton : ITaiwanCalendarSingleton
    {
        public TaiwanCalendarSingleton() { }

        // Factory Methods
        public TaiwanCalendar Create()
        {
            return new TaiwanCalendar();
        }

        // Fields
        public int CurrentEra => GGlobalization.Calendar.CurrentEra;


        // Methods
        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.TaiwanCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

        public Calendar ReadOnly(Calendar pCalendar)
        {
            return GGlobalization.Calendar.ReadOnly((GGlobalization.Calendar)pCalendar.Unwrap()).Wrap();
        }

    }
}
