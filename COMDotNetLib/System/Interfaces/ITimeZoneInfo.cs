using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNetLib.System
{
    public interface ITimeZoneInfo
    {
        // Properties

        ITimeSpan BaseUtcOffset
        {
            [Description("Gets the time difference between the current time zone's standard time and Coordinated Universal Time (UTC).")]
            get;
        }

        string DaylightName
        {
            [Description("Gets the display name for the current time zone's daylight saving time.")]
            get;
        }

        string DisplayName
        {
            [Description("Gets the general display name that represents the time zone.")]
            get;
        }

        string Id
        {
            [Description("Gets the time zone identifier.")]
            get;
        }

        string StandardName
        {
            [Description("Gets the display name for the time zone's standard time.")]
            get;
        }

        bool SupportsDaylightSavingTime
        {
            [Description("Gets a value indicating whether the time zone has any daylight saving time rules.")]
            get;
        }

        // Methods

        [Description("Determines whether the current TimeZoneInfo object and another TimeZoneInfo object are equal.")]
        bool Equals(TimeZoneInfo other);

        [Description("Determines whether the current TimeZoneInfo object and another object are equal.")]
        bool Equals2(object obj);

        [Description("Returns information about the possible dates and times that an ambiguous date and time can be mapped to.")]
        ITimeSpan[] GetAmbiguousTimeOffsets(DateTime dateTime);

        [Description("Returns information about the possible dates and times that an ambiguous date and time can be mapped to.")]
        ITimeSpan[] GetAmbiguousTimeOffsets2(DateTimeOffset dateTimeOffset);

        [Description("Serves as a hash function for hashing algorithms and data structures such as hash tables.")]
        int GetHashCode();

        [Description("Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time.")]
        ITimeSpan GetUtcOffset(DateTime dateTime);

        [Description("Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time.")]
        ITimeSpan GetUtcOffset2(DateTimeOffset dateTimeOffset);

        [Description("Indicates whether the current object and another TimeZoneInfo object have the same adjustment rules.")]
        bool HasSameRules(TimeZoneInfo other);

        [Description("Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times.")]
        bool IsAmbiguousTime(DateTime dateTime);

        [Description("Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times.")]
        bool IsAmbiguousTime2(DateTimeOffset dateTimeOffset);

        [Description("Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object.")]
        bool IsDaylightSavingTime(DateTime dateTime);

        [Description("Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object.")]
        bool IsDaylightSavingTime2(DateTimeOffset dateTimeOffset);

        [Description("Indicates whether a particular date and time is invalid.")]
        bool IsInvalidTime(DateTime dateTime);

        [Description("Converts the current TimeZoneInfo object to a serialized string.")]
        string ToSerializedString();
        string ToString();
    }
}
