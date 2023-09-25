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
    [Guid("8C15E7B5-9D80-46E1-9EFB-430A95A235AF")]
    [Description("Represents the Julian calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IJulianCalendarSingleton : ICalendarSingleton
    {
        [Description("Initializes a new instance of the JulianCalendar class.")]
        JulianCalendar Create();

        // Fields
        new int CurrentEra
        {
            [Description("Represents the current era of the current calendar. The value of this field is 0.")]
            get;
        }

        int JulianEra
        {
            [Description("Represents the current era. This field is constant.")]
            get;
        }


        // Methods
        [Description("Determines whether the specified object instances are considered equal.")]
        new bool Equals(object objA, object objB);

        [Description("Returns a read-only version of the specified Calendar object.")]
        new Calendar ReadOnly(Calendar pCalendar);
    }
}
