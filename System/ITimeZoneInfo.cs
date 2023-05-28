using DotNetLib.System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    // https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8.1

    [ComVisible(true)]
    [Guid("3484FC28-42F0-4BCA-A5C4-F7AF3DD4D441")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITimeZoneInfo
    {
        //Description("")]

        // Properties

        [Description("Gets the time difference between the current time zone's standard time and Coordinated Universal Time (UTC).")]
        TimeSpan BaseUtcOffset { get; }

        [Description("Gets the display name for the current time zone's daylight saving time.")]
        string DaylightName { get; }

        [Description("Gets the general display name that represents the time zone.")]
        string DisplayName { get; }
        
        [Description("Gets the time zone identifier.")]
        string Id { get; }

        [Description("Gets a TimeZoneInfo object that represents the local time zone.")]
        TimeZoneInfo Local { get; }

        [Description("Gets the display name for the time zone's standard time.")]
        string StandardName { get; }

        [Description("Gets a value indicating whether the time zone has any daylight saving time rules.")]
        bool SupportsDaylightSavingTime { get; }
        
        [Description("Gets a TimeZoneInfo object that represents the Coordinated Universal Time (UTC) zone.")]
        TimeZoneInfo Utc { get; }

        // Methods

        [Description("Clears cached time zone data.")]
        void ClearCachedData();

        [Description("Converts a time to the time in a particular time zone.")]
        DateTime ConvertTime(DateTime dateTime, TimeZoneInfo destinationTimeZone);

        [Description("Converts a time to the time in a particular time zone.")]
        DateTimeOffset ConvertTime2(DateTimeOffset dateTimeOffset, TimeZoneInfo destinationTimeZone);

        [Description("Converts a time from one time zone to another.")]
        DateTime ConvertTime3(DateTime dateTime, TimeZoneInfo sourceTimeZone, TimeZoneInfo destinationTimeZone);

        [Description("Converts a time to the time in another time zone based on the time zone's identifier.")]
        DateTime ConvertTimeBySystemTimeZoneId(DateTime dateTime, string destinationTimeZoneId);

        [Description("Converts a time to the time in another time zone based on the time zone's identifier.")]
        DateTimeOffset ConvertTimeBySystemTimeZoneId2(DateTimeOffset dateTimeOffset, string destinationTimeZoneId);

        [Description("Converts a time from one time zone to another based on time zone identifiers.")]
        DateTime ConvertTimeBySystemTimeZoneId3(DateTime dateTime, string sourceTimeZoneId, string destinationTimeZoneId);


        [Description("Converts a Coordinated Universal Time (UTC) to the time in a specified time zone.")]
        DateTime ConvertTimeFromUtc(DateTime dateTime, TimeZoneInfo destinationTimeZone);

        [Description("Converts the specified date and time to Coordinated Universal Time (UTC).")]
        DateTime ConvertTimeToUtc(DateTime dateTime);

        [Description("Converts the time in a specified time zone to Coordinated Universal Time (UTC).")]
        DateTime ConvertTimeToUtc2(DateTime dateTime, TimeZoneInfo sourceTimeZone);

        [Description("Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, and a standard time display name.")]
        TimeZoneInfo CreateCustomTimeZone(string id, TimeSpan baseUtcOffset, string displayName, string standardDisplayName);

        //[Description("Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, a standard time name, a daylight saving time name, and daylight saving time rules.")]
        //TODO: TimeZoneInfo CreateCustomTimeZone2(string id, TimeSpan baseUtcOffset, string displayName, string standardDisplayName, string daylightDisplayName, TimeZoneInfo.AdjustmentRule[] adjustmentRules);

        //[Description("Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, a standard time name, a daylight saving time name, daylight saving time rules, and a value that indicates whether the returned object reflects daylight saving time information.")]
        //TODO: TimeZoneInfo CreateCustomTimeZone3(string id, TimeSpan baseUtcOffset, string displayName, string standardDisplayName, string daylightDisplayName, TimeZoneInfo.AdjustmentRule[] adjustmentRules, bool disableDaylightSavingTime);

        [Description("Determines whether the current TimeZoneInfo object and another TimeZoneInfo object are equal.")]
        bool Equals(TimeZoneInfo other);

        [Description("Determines whether the current TimeZoneInfo object and another object are equal.")]
        bool Equals2(object obj);

        [Description("Instantiates a new TimeZoneInfo object based on its identifier.")]
        TimeZoneInfo FindSystemTimeZoneById(string id);


        [Description("Deserializes a string to re-create an original serialized TimeZoneInfo object.")]
        TimeZoneInfo FromSerializedString(string source);

        //[Description("Retrieves an array of TimeZoneInfo.AdjustmentRule objects that apply to the current TimeZoneInfo object.")]
        //TimeZoneInfo.AdjustmentRule[] GetAdjustmentRules();

        [Description("Returns information about the possible dates and times that an ambiguous date and time can be mapped to.")]
        TimeSpan[] GetAmbiguousTimeOffsets(DateTime dateTime);

        [Description("Returns information about the possible dates and times that an ambiguous date and time can be mapped to.")]
        TimeSpan[] GetAmbiguousTimeOffsets2(DateTimeOffset dateTimeOffset);

        [Description("Serves as a hash function for hashing algorithms and data structures such as hash tables.")]
        int GetHashCode();

        [Description("Returns a sorted collection of all the time zones about which information is available on the local system.")]
        ReadOnlyCollection GetSystemTimeZones();

        [Description("Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time.")] 
        TimeSpan GetUtcOffset(DateTime dateTime);

        [Description("Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time.")] 
        TimeSpan GetUtcOffset2(DateTimeOffset dateTimeOffset);

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
