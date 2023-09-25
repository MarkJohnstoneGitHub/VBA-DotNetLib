// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.thaibuddhistcalendar?view=netframework-4.8.1

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
    [Guid("5ADB3229-B9EA-45C2-A8E3-CE9AFA36D6E4")]
    [ProgId("DotNetLib.System.Globalization.ThaiBuddhistCalendarSingleton")]
    [Description("Represents the Thai Buddhist calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IThaiBuddhistCalendarSingleton))]
    public class ThaiBuddhistCalendarSingleton : IThaiBuddhistCalendarSingleton
    {
        public ThaiBuddhistCalendarSingleton() {}

        // Factory Methods
        public ThaiBuddhistCalendar Create()
        {
            return new ThaiBuddhistCalendar();
        }

        // Fields
        public int CurrentEra => GGlobalization.Calendar.CurrentEra;

        public int ThaiBuddhistEra => GGlobalization.ThaiBuddhistCalendar.ThaiBuddhistEra;


        // Methods
        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.UmAlQuraCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

        public Calendar ReadOnly(Calendar pCalendar)
        {
            return GGlobalization.Calendar.ReadOnly((GGlobalization.Calendar)pCalendar.Unwrap()).Wrap();
        }

    }
}
