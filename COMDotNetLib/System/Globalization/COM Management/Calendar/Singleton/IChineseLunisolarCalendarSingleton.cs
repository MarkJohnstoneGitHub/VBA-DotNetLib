//  https://learn.microsoft.com/en-us/dotnet/api/system.globalization.chineselunisolarcalendar?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("098ABEF8-2092-46BE-ACDA-628F1E128200")]
    [Description("Represents time in divisions, such as months, days, and years. Years are calculated using the Chinese calendar, while days and months are calculated using the lunisolar calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IChineseLunisolarCalendarSingleton : ICalendarSingleton
    {
        // Factory Methods

        [Description("Initializes a new instance of the ChineseLunisolarCalendar class.")]
        ChineseLunisolarCalendar Create();

        // Fields

        new int CurrentEra
        {
            [Description("Represents the current era of the current calendar. The value of this field is 0.")]
            get;
        }

        int ChineseEra
        {
            [Description("Specifies the era that corresponds to the current ChineseLunisolarCalendar object.")]
            get;
        }

        // Methods

        [Description("Determines whether the specified object instances are considered equal.")]
        new bool Equals(object objA, object objB);

        [Description("Returns a read-only version of the specified Calendar object.")]
        new ICalendar ReadOnly(ICalendar pCalendar);
    }
}
