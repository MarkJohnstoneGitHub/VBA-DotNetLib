// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.hijricalendar?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using DotNetLib.Extensions;
using System.Globalization;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("7F5909B3-125E-45C4-A9A5-AF68E73318F8")]
    [ProgId("DotNetLib.System.Globalization.HijriCalendarSingleton")]
    [Description("Represents the Hijri calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IHijriCalendarSingleton))]
    public class HijriCalendarSingleton : IHijriCalendarSingleton
    {
        public HijriCalendarSingleton() { }

        public HijriCalendar Create()
        {
            return new HijriCalendar();
        }

        // Fields
        public int CurrentEra => Calendar.CurrentEra;

        public int HijriEra => GGlobalization.HijriCalendar.HijriEra;

        // Methods
        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.HijriCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

        public ICalendar ReadOnly(ICalendar calendar)
        {
            return Calendar.ReadOnly((Calendar)calendar.Unwrap()).Wrap();
        }

    }
}
