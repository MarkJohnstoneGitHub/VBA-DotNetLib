// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.umalquracalendar?view=netframework-4.8.1

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
    [Guid("5B4D196C-CBBD-4AC2-8A4C-4EECC5DC9BC2")]
    [ProgId("DotNetLib.System.Globalization.UmAlQuraCalendarSingleton")]
    [Description("Represents the UmAlQuraCalendar calendar factory methods and static members.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IUmAlQuraCalendarSingleton))]
    public class UmAlQuraCalendarSingleton : IUmAlQuraCalendarSingleton
    {
        public UmAlQuraCalendarSingleton() {}

        // Factory Methods

        public UmAlQuraCalendar Create()
        {
            return new UmAlQuraCalendar();
        }

        // Fields
        public int CurrentEra => Calendar.CurrentEra;

        public int UmAlQuraEra => GGlobalization.UmAlQuraCalendar.UmAlQuraEra;


        // Methods
        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.UmAlQuraCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

        public ICalendar ReadOnly(ICalendar calendar)
        {
            return Calendar.ReadOnly((Calendar)calendar.Unwrap()).Wrap();
        }


    }
}
