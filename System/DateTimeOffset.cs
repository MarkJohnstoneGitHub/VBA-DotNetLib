using GSystem = global::System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;

namespace DotNetLib.System
{
    // https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1

    [ComVisible(true)]
    [Description("Represents a point in time, typically expressed as a date and time of day, relative to Coordinated Universal Time (UTC).")]
    [Guid("27660912-6101-4779-B4E0-24F2B164B334")]
    [ProgId("DotNetLib.System.DateTimeOffset")]
    [ClassInterface(ClassInterfaceType.None)]
    public class DateTimeOffset //: IDateTimeOffset
    {
        private GSystem.DateTimeOffset objDateTimeOffset;

        public static DateTimeOffset dtoMaxValue = new DateTimeOffset(GSystem.DateTimeOffset.MaxValue);
        public static readonly DateTimeOffset dtoMinValue = new DateTimeOffset(GSystem.DateTimeOffset.MinValue);

        // Constructors

        internal DateTimeOffset(GSystem.DateTimeOffset objDateTimeOffset)
        {
            this.objDateTimeOffset = objDateTimeOffset;
        }

        public DateTimeOffset()
        {
            objDateTimeOffset = new GSystem.DateTimeOffset();
        }

        public DateTimeOffset(DateTime dateTime)
        {
            objDateTimeOffset = new GSystem.DateTimeOffset(dateTime.dateTime);
        }

        public DateTimeOffset(DateTime dateTime, TimeSpan offset)
        {
            objDateTimeOffset = new GSystem.DateTimeOffset(dateTime.dateTime,offset.timeSpan);
        }

        public DateTimeOffset(long ticks, TimeSpan offset)
        {
            objDateTimeOffset = new GSystem.DateTimeOffset(ticks, offset.timeSpan);
        }

        public DateTimeOffset(int year, int month, int day, int hour, int minute, int second, TimeSpan offset)
        {
            objDateTimeOffset = new GSystem.DateTimeOffset(year,month,day,hour,minute,second, offset.timeSpan);
        }

        public DateTimeOffset(int year, int month, int day, int hour, int minute, int second, int millisecond, TimeSpan offset)
        {
            objDateTimeOffset = new GSystem.DateTimeOffset(year, month, day, hour, minute, second, millisecond, offset.timeSpan);
        }

        public DateTimeOffset(int year, int month, int day, int hour, int minute, int second, int millisecond, GSystem.Globalization.Calendar calendar, TimeSpan offset)
        {
            objDateTimeOffset = new GSystem.DateTimeOffset(year, month, day, hour, minute, second, millisecond, calendar, offset.timeSpan);
        }

        internal GSystem.DateTimeOffset dateTimeOffset
        {
            get { return this.objDateTimeOffset; }
        }

        public DateTimeOffset CreateFromDateTime(DateTime dateTime)
        {
            return new DateTimeOffset(dateTime);
        }

        public DateTimeOffset CreateFromDateTime2(DateTime dateTime, TimeSpan offset)
        {
            return new DateTimeOffset(dateTime,offset);
        }

        public DateTimeOffset CreateFromDateTimeParts(int year, int month, int day, int hour, int minute, int second, TimeSpan offset)
        {
            return new DateTimeOffset(year, month, day, hour, minute, second, offset);
        }

        public DateTimeOffset CreateFromDateTimeParts2(int year, int month, int day, int hour, int minute, int second, int millisecond, TimeSpan offset)
        {
            return new DateTimeOffset(year, month, day, hour, minute, second, millisecond, offset);
        }

        public DateTimeOffset CreateFromDateTimeParts3(int year, int month, int day, int hour, int minute, int second, int millisecond, GSystem.Globalization.Calendar calendar, TimeSpan offset)
        {
            return new DateTimeOffset(year, month, day, hour, minute, second, millisecond, calendar, offset);
        }

        public DateTimeOffset CreateFromTicks(long ticks, TimeSpan offset)
        {
            return new DateTimeOffset(ticks, offset);
        }

        // Fields
        public DateTimeOffset MaxValue => dtoMaxValue;
        public DateTimeOffset MinValue => dtoMinValue;

        public DateTime Date
        {
            get {return new DateTime(this.objDateTimeOffset.Date); }
        }

        public DateTime DateTime
        {
            get { return new DateTime(this.objDateTimeOffset.DateTime); }
        }
        
        public int Day => this.objDateTimeOffset.Day;

        public DayOfWeek DayOfWeek => (DayOfWeek)this.objDateTimeOffset.DayOfWeek;


        public int DayOfYear => this.objDateTimeOffset.DayOfYear;

        public int Hour => this.objDateTimeOffset.Hour;

        public DateTime LocalDateTime
        {
            get { return new DateTime(this.objDateTimeOffset.LocalDateTime); }
        }
        
        public int Millisecond => this.objDateTimeOffset.Millisecond;

        public int Minute => this.objDateTimeOffset.Minute;

        public int Month => this.objDateTimeOffset.Month;

        public DateTimeOffset Now
        {
            get { return new DateTimeOffset(GSystem.DateTimeOffset.Now); }
        }

        public TimeSpan Offset
        {
            get { return new TimeSpan(this.objDateTimeOffset.Offset); }
        }
        
        public int Second => this.objDateTimeOffset.Second;

        public long Ticks => this.objDateTimeOffset.Ticks;

        public TimeSpan TimeOfDay
        {
            get { return new TimeSpan(this.objDateTimeOffset.TimeOfDay); }
        }

        public DateTime UtcDateTime
        {
            get { return new DateTime(this.objDateTimeOffset.UtcDateTime); }
        }        
        
        public DateTimeOffset UtcNow
        {
            get { return new DateTimeOffset(GSystem.DateTimeOffset.UtcNow); }
        }   

        public long UtcTicks => this.objDateTimeOffset.UtcTicks;

        public int Year => this.objDateTimeOffset.Year;

        // Methods

        public DateTimeOffset Add(TimeSpan timeSpan)
        {
            return new DateTimeOffset(this.objDateTimeOffset.Add(timeSpan.timeSpan));
        }

        public DateTimeOffset AddDays(double days)
        {
            return new DateTimeOffset(this.objDateTimeOffset.AddDays(days));
        }

    }


}
