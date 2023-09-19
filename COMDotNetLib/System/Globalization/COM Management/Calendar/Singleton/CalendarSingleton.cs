// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.calendar?view=netframework-4.8.1
// Notes: https://stackoverflow.com/questions/19075759/convert-between-calendars

using GGlobalization = global::System.Globalization;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.Extensions;
using System.Globalization;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("B6635459-034A-4D2A-A95A-7EB5A280E243")]
    [ProgId("DotNetLib.System.Globalization.CalendarSingleton")]
    [Description("Represents the calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICalendarSingleton))]
    public class CalendarSingleton : ICalendarSingleton
    {
        public CalendarSingleton() { }

        // Fields
        public int CurrentEra => GGlobalization.Calendar.CurrentEra;

        // Methods
        public new bool Equals(object objA, object objB)
        {  
            return GGlobalization.Calendar.Equals(objA.Unwrap(), objB.Unwrap()); 
        }

        public ICalendar ReadOnly(ICalendar calendar)
        {
            return Calendar.ReadOnly((Calendar)calendar.Unwrap()).Wrap();
        }

    }
}
