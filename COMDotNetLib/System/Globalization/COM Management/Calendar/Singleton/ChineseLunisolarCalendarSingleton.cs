//  https://learn.microsoft.com/en-us/dotnet/api/system.globalization.chineselunisolarcalendar?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using DotNetLib.Extensions;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("5943D123-799C-4C3A-91AB-5588D692E357")]
    [ProgId("DotNetLib.System.Globalization.ChineseLunisolarCalendarSingleton")]
    [Description("Represents time in divisions, such as months, days, and years. Years are calculated using the Chinese calendar, while days and months are calculated using the lunisolar calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IChineseLunisolarCalendarSingleton))]
    public class ChineseLunisolarCalendarSingleton : IChineseLunisolarCalendarSingleton
    {
        public ChineseLunisolarCalendarSingleton() { }

        // Factory Methods
        public ChineseLunisolarCalendar Create()
        {
            return new ChineseLunisolarCalendar();
        }

        // Fields
        public int CurrentEra => Calendar.CurrentEra;

        public int ChineseEra => GGlobalization.ChineseLunisolarCalendar.ChineseEra;

        // Methods
        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.HebrewCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

        public ICalendar ReadOnly(ICalendar pCalendar)
        {
            return Calendar.ReadOnly((Calendar)pCalendar.Unwrap()).Wrap();
        }

    }
}
