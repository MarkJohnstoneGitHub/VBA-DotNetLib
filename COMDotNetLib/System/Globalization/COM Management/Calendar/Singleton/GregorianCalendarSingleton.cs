using GGlobalization = global::System.Globalization;
using DotNetLib.Extensions;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("69AFAFEA-385A-48CE-91C0-790459722718")]
    [ProgId("DotNetLib.System.Globalization.GregorianCalendarSingleton")]
    [Description("Represents the Gregorian calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IGregorianCalendarSingleton))]
    public class GregorianCalendarSingleton : IGregorianCalendarSingleton
    {
        public GregorianCalendarSingleton() { }


        // Factory Methods

        public GregorianCalendar Create(GregorianCalendarTypes type = GregorianCalendarTypes.Localized)
        {
            return new GregorianCalendar(type);
        }

        // Fields

        public int CurrentEra => Calendar.CurrentEra;

        public int ADEra => GGlobalization.GregorianCalendar.ADEra;


        public ICalendar ReadOnly(ICalendar calendar)
        {
            //Calendar calendarUnwrapped = (Calendar)calendar.Unwrap();
            //Calendar calendarReadOnly = Calendar.ReadOnly(calendarUnwrapped);
            //return calendarReadOnly.Wrap();

            return Calendar.ReadOnly((Calendar)calendar.Unwrap()).Wrap();
        }

        // Methods

        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.GregorianCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

    }
}
