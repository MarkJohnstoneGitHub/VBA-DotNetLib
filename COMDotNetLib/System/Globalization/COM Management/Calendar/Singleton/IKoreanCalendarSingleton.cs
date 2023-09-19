// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.koreancalendar?view=netframework-4.8.1

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("E4D548E6-F183-4D5D-854E-5E785BDED14E")]
    [Description("Represents the Korean calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IKoreanCalendarSingleton : ICalendarSingleton
    {
        // Factory Methods

        [Description("Initializes a new instance of the HebrewCalendar class.")]
        KoreanCalendar Create();

        // Fields

        new int CurrentEra
        {
            [Description("Represents the current era of the current calendar. The value of this field is 0.")]
            get;
        }

        int KoreanEra
        {
            [Description("Represents the current era. This field is constant.")]
            get;
        }

        // Methods

        [Description("Determines whether the specified object instances are considered equal.")]
        new bool Equals(object objA, object objB);

        [Description("Returns a read-only version of the specified Calendar object.")]
        new ICalendar ReadOnly(ICalendar calendar);

    }
}
