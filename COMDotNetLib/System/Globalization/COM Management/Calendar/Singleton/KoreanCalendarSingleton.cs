// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.koreancalendar?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using DotNetLib.Extensions;
using System;
using System.Globalization;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("1752AC9A-2104-4BB2-8DF9-B66FB65412DB")]
    [ProgId("DotNetLib.System.Globalization.KoreanCalendarSingleton")]
    [Description("Represents the Korean calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IKoreanCalendarSingleton))]
    public class KoreanCalendarSingleton : IKoreanCalendarSingleton
    {
        public KoreanCalendarSingleton() { }

        // Factory Methods
        public KoreanCalendar Create()
        {
            return new KoreanCalendar();
        }

        //Fields
        public int CurrentEra => Calendar.CurrentEra;

        public int KoreanEra => GGlobalization.KoreanCalendar.KoreanEra;

        // Methods
        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.HebrewCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

        public ICalendar ReadOnly(ICalendar calendar)
        {
            return Calendar.ReadOnly((Calendar)calendar.Unwrap()).Wrap();
        }

    }
}
