// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.japanesecalendar?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.Extensions;
using System.Globalization;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("63EA1611-1E33-4BF5-B1EB-CDFED88A36F3")]
    [ProgId("DotNetLib.System.Globalization.JapaneseCalendarSingleton")]
    [Description("Represents the Japanese calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IJapaneseCalendarSingleton))]
    public class JapaneseCalendarSingleton : IJapaneseCalendarSingleton
    {
        public JapaneseCalendarSingleton() {}

        public JapaneseCalendar Create()
        {
            return new JapaneseCalendar();
        }

        // Fields

        public int CurrentEra => Calendar.CurrentEra;

        // Methods
        
        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.JapaneseCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

        public ICalendar ReadOnly(ICalendar calendar)
        {
            return Calendar.ReadOnly((Calendar)calendar.Unwrap()).Wrap();
        }
    }
}
