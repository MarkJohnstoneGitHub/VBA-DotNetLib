using GSystem = global::System; // https://stackoverflow.com/questions/5681537/namespace-conflict-in-c-sharp
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("818900B7-0353-45FA-8038-1C550219FD04")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]

    public interface IDateTimeOffset
    {
        // Constructors

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified DateTime value.")]
        DateTimeOffset CreateFromDateTime(DateTime dateTime);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified DateTime value and offset.")]
        DateTimeOffset CreateFromDateTime2(DateTime dateTime, TimeSpan offset);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, and offset.")]
        DateTimeOffset CreateFromDateTimeParts(int year, int month, int day, int hour, int minute, int second, TimeSpan offset);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, millisecond, and offset.")]
        DateTimeOffset CreateFromDateTimeParts2(int year, int month, int day, int hour, int minute, int second, int millisecond, TimeSpan offset);

        DateTimeOffset CreateFromDateTimeParts3(int year, int month, int day, int hour, int minute, int second, int millisecond, GSystem.Globalization.Calendar calendar, TimeSpan offset);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified number of ticks and offset.")]
        DateTimeOffset CreateFromTicks(long ticks, TimeSpan offset);


        // Fields

        [Description("Represents the greatest possible value of DateTimeOffset. This field is read-only.")]
        DateTimeOffset MaxValue { get; }

        [Description("Represents the earliest possible DateTimeOffset value. This field is read-only.")]
        DateTimeOffset MinValue { get; }

        // Properties

        [Description("Gets a DateTime value that represents the date component of the current DateTimeOffset object.")]
        DateTime Date { get; }

        [Description("Gets a DateTime value that represents the date and time of the current DateTimeOffset object.")]
        DateTime DateTime { get; }

        [Description("Gets the day of the month represented by the current DateTimeOffset object.")]
        int Day { get; }

        [Description("Gets the day of the week represented by the current DateTimeOffset object.")]
        DayOfWeek DayOfWeek { get; }

        [Description("Gets the day of the year represented by the current DateTimeOffset object.")]
        int DayOfYear { get; }

        [Description("Gets the hour component of the time represented by the current DateTimeOffset object.")]
        int Hour { get; }

        [Description("Gets a DateTime value that represents the local date and time of the current DateTimeOffset object.")]
        DateTime LocalDateTime { get; }

        [Description("Gets the millisecond component of the time represented by the current DateTimeOffset object.")]
        int Millisecond { get; }

        [Description("Gets the minute component of the time represented by the current DateTimeOffset object.")]
        int Minute { get; }

        [Description("Gets the month component of the date represented by the current DateTimeOffset object.")]
        int Month { get; }

        [Description("Gets a DateTimeOffset object that is set to the current date and time on the current computer, with the offset set to the local time's offset from Coordinated Universal Time (UTC).")]
        DateTimeOffset Now { get; }

        [Description("Gets the time's offset from Coordinated Universal Time (UTC).")]
        TimeSpan Offset { get; }

        [Description("Gets the second component of the clock time represented by the current DateTimeOffset object.")]
        int Second { get; }

        [Description("Gets the number of ticks that represents the date and time of the current DateTimeOffset object in clock time.")]
        long Ticks { get; }

        [Description("Gets the time of day for the current DateTimeOffset object.")]
        TimeSpan TimeOfDay { get; }

        [Description("Gets a DateTime value that represents the Coordinated Universal Time (UTC) date and time of the current DateTimeOffset object.")]
        DateTime UtcDateTime { get; }

        [Description("Gets a DateTimeOffset object whose date and time are set to the current Coordinated Universal Time (UTC) date and time and whose offset is Zero.")]
        DateTimeOffset UtcNow { get; }

        [Description("Gets the number of ticks that represents the date and time of the current DateTimeOffset object in Coordinated Universal Time (UTC).")]
        long UtcTicks { get; }

        [Description("Gets the year component of the date represented by the current DateTimeOffset object.")]
        int Year { get; }

        // Methods

        [Description("Returns a new DateTimeOffset object that adds a specified time interval to the value of this instance.")]
        DateTimeOffset Add(TimeSpan timeSpan);

        [Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional days to the value of this instance.")]
        DateTimeOffset AddDays(double days);

        [Description("Returns a new DateTimeOffset object that adds a specified number of milliseconds to the value of this instance.")]
        DateTimeOffset AddMilliseconds(double milliseconds);

        DateTimeOffset AddMinutes(double minutes);

        //[Description("")]
    }
}
