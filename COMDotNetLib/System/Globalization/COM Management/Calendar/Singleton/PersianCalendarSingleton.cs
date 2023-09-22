// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.persiancalendar?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using DotNetLib.Extensions;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("96A9C220-CC8D-4EE5-BE70-675AB27EBC81")]
    [ProgId("DotNetLib.System.Globalization.PersianCalendarSingleton")]
    [Description("Represents the Persian calendar factory methods and static members.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IPersianCalendarSingleton))]
    public class PersianCalendarSingleton : IPersianCalendarSingleton
    {
        public PersianCalendarSingleton() {}

        public PersianCalendar Create()
        {
            return new PersianCalendar();
        }

        // Fields
        public int CurrentEra => GGlobalization.Calendar.CurrentEra;

        public int PersianEra => GGlobalization.PersianCalendar.PersianEra;

        // Methods
        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.PersianCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

        public Calendar ReadOnly(Calendar pCalendar)
        {
            return GGlobalization.Calendar.ReadOnly((GGlobalization.Calendar)pCalendar.Unwrap()).Wrap();
        }

    }
}
