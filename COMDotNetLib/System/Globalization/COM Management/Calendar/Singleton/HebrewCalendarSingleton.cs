// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.hebrewcalendar?view=netframework-4.8.1

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
    [Guid("99057C8A-82CA-4C12-B447-4D276820F913")]
    [ProgId("DotNetLib.System.Globalization.HebrewCalendarSingleton")]
    [Description("Represents the Hebrew calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IHebrewCalendarSingleton))]
    public class HebrewCalendarSingleton : IHebrewCalendarSingleton
    {
        public HebrewCalendarSingleton() { }

        // Factory Methods
        public HebrewCalendar Create()
        {
            return new HebrewCalendar();
        }

        // Fields
        public int CurrentEra => GGlobalization.Calendar.CurrentEra;

        public int HebrewEra => GGlobalization.HebrewCalendar.HebrewEra;

        // Methods
        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.HebrewCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

        public Calendar ReadOnly(Calendar pCalendar)
        {
            return GGlobalization.Calendar.ReadOnly((GGlobalization.Calendar)pCalendar.Unwrap()).Wrap();
        }

    }
}
