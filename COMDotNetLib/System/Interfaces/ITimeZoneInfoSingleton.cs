// https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1

using DotNetLib.System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("3484FC28-42F0-4BCA-A5C4-F7AF3DD4D441")]
    [TypeLibTypeAttribute(TypeLibTypeFlags.FPreDeclId)] //The type is predefined. The client application should automatically create a single instance of the object that has this attribute. 
    [Description("Represents any time zone in the world.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITimeZoneInfoSingleton 
    {
        // Properties

        ITimeZoneInfo Locale
        {
            [Description("Gets a TimeZoneInfo object that represents the local time zone.")]
            get;
        }

        ITimeZoneInfo Utc
        {
            [Description("Gets a TimeZoneInfo object that represents the Coordinated Universal Time (UTC) zone.")]
            get;
        }

        // Methods

        [Description("Clears cached time zone data.")]
        void ClearCachedData();

        [Description("Converts a time to the time in a particular time zone.")]
        IDateTime ConvertTime(DateTime dateTime, TimeZoneInfo destinationTimeZone);

        [Description("Converts a time to the time in a particular time zone.")]
        IDateTimeOffset ConvertTime2(DateTimeOffset dateTimeOffset, TimeZoneInfo destinationTimeZone);

        [Description("Converts a time from one time zone to another.")]
        IDateTime ConvertTime3(DateTime dateTime, TimeZoneInfo sourceTimeZone, TimeZoneInfo destinationTimeZone);

        [Description("Converts a time to the time in another time zone based on the time zone's identifier.")]
        IDateTime ConvertTimeBySystemTimeZoneId(DateTime dateTime, string destinationTimeZoneId);

        [Description("Converts a time to the time in another time zone based on the time zone's identifier.")]
        IDateTimeOffset ConvertTimeBySystemTimeZoneId2(DateTimeOffset dateTimeOffset, string destinationTimeZoneId);

        [Description("Converts a time from one time zone to another based on time zone identifiers.")]
        IDateTime ConvertTimeBySystemTimeZoneId3(DateTime dateTime, string sourceTimeZoneId, string destinationTimeZoneId);

        [Description("Converts a Coordinated Universal Time (UTC) to the time in a specified time zone.")]
        IDateTime ConvertTimeFromUtc(DateTime dateTime, TimeZoneInfo destinationTimeZone);

        [Description("Converts the specified date and time to Coordinated Universal Time (UTC).")]
        IDateTime ConvertTimeToUtc(DateTime dateTime);

        [Description("Converts the time in a specified time zone to Coordinated Universal Time (UTC).")]
        IDateTime ConvertTimeToUtc2(DateTime dateTime, TimeZoneInfo sourceTimeZone);

        [Description("Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, and a standard time display name.")]
        ITimeZoneInfo CreateCustomTimeZone(string pId, TimeSpan pBaseUtcOffset, string pDisplayName, string standardDisplayName);

        //[Description("Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, a standard time name, a daylight saving time name, and daylight saving time rules.")]
        //TODO: TimeZoneInfo CreateCustomTimeZone2(string id, TimeSpan baseUtcOffset, string displayName, string standardDisplayName, string daylightDisplayName, TimeZoneInfo.AdjustmentRule[] adjustmentRules);

        //[Description("Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, a standard time name, a daylight saving time name, daylight saving time rules, and a value that indicates whether the returned object reflects daylight saving time information.")]
        //TODO: TimeZoneInfo CreateCustomTimeZone3(string id, TimeSpan baseUtcOffset, string displayName, string standardDisplayName, string daylightDisplayName, TimeZoneInfo.AdjustmentRule[] adjustmentRules, bool disableDaylightSavingTime);


        [Description("Instantiates a new TimeZoneInfo object based on its identifier.")]
        ITimeZoneInfo FindSystemTimeZoneById(string pId);

        [Description("Deserializes a string to re-create an original serialized TimeZoneInfo object.")]
        ITimeZoneInfo FromSerializedString(string source);

        //[Description("Retrieves an array of TimeZoneInfo.AdjustmentRule objects that apply to the current TimeZoneInfo object.")]
        //TODO: TimeZoneInfo.AdjustmentRule[] GetAdjustmentRules();

        [Description("Returns a sorted collection of all the time zones about which information is available on the local system.")]
        ReadOnlyCollection GetSystemTimeZones();
    }
}
