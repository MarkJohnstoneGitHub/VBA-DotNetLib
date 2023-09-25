// https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1

using DotNetLib.System.Collections;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Description("TimeZoneInfo factory methods and static members.")]
    [Guid("0EACC4BC-94EF-45D6-8FB8-5A90E119EA16")]
    //TypeLibTypeAttribute(TypeLibTypeFlags.FPreDeclId)] //The type is predefined. The client application should automatically create a single instance of the object that has this attribute. 
    [TypeLibType(TypeLibTypeFlags.FPreDeclId | TypeLibTypeFlags.FCanCreate)] //The type is predefined. The client application should automatically create a single instance of the object that has this attribute. 
    [ProgId("DotNetLib.System.TimeZoneInfo")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITimeZoneInfoSingleton))]
    public class TimeZoneInfoSingleton : ITimeZoneInfoSingleton
    {
        public TimeZoneInfoSingleton() { }

        // Properties
        public TimeZoneInfo Local => TimeZoneInfo.Local;

        public TimeZoneInfo Utc => TimeZoneInfo.Utc;

        // Methods

        public void ClearCachedData()
        {
            TimeZoneInfo.ClearCachedData();
        }

        public DateTime ConvertTime(DateTime dateTime, TimeZoneInfo destinationTimeZone)
        {
            return TimeZoneInfo.ConvertTime(dateTime, destinationTimeZone);
        }

        public DateTimeOffset ConvertTime2(DateTimeOffset dateTimeOffset, TimeZoneInfo destinationTimeZone)
        {
            return TimeZoneInfo.ConvertTime(dateTimeOffset, destinationTimeZone);
        }

        public DateTime ConvertTime3(DateTime dateTime, TimeZoneInfo sourceTimeZone, TimeZoneInfo destinationTimeZone)
        {
            return TimeZoneInfo.ConvertTime(dateTime, sourceTimeZone, destinationTimeZone);
        }

        public DateTime ConvertTimeBySystemTimeZoneId(DateTime dateTime, string destinationTimeZoneId)
        {
            return TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime, destinationTimeZoneId);
        }

        public DateTimeOffset ConvertTimeBySystemTimeZoneId2(DateTimeOffset dateTimeOffset, string destinationTimeZoneId)
        {
            return TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTimeOffset, destinationTimeZoneId);
        }

        public DateTime ConvertTimeBySystemTimeZoneId3(DateTime dateTime, string sourceTimeZoneId, string destinationTimeZoneId)
        {
            return TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime, sourceTimeZoneId, destinationTimeZoneId);
        }

        public DateTime ConvertTimeFromUtc(DateTime dateTime, TimeZoneInfo destinationTimeZone)
        {
            return TimeZoneInfo.ConvertTimeFromUtc(dateTime, destinationTimeZone);
        }

        public DateTime ConvertTimeToUtc(DateTime dateTime)
        {
            return TimeZoneInfo.ConvertTimeToUtc(dateTime);
        }
        public DateTime ConvertTimeToUtc2(DateTime dateTime, TimeZoneInfo sourceTimeZone)
        {
            return TimeZoneInfo.ConvertTimeToUtc(dateTime, sourceTimeZone);
        }

        public TimeZoneInfo CreateCustomTimeZone(string pId, TimeSpan pBaseUtcOffset, string pDisplayName, string standardDisplayName)
        {
            return TimeZoneInfo.CreateCustomTimeZone(pId, pBaseUtcOffset, pDisplayName, standardDisplayName);
        }

        //[Description("Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, a standard time name, a daylight saving time name, and daylight saving time rules.")]
        //TODO: TimeZoneInfo CreateCustomTimeZone2(string id, TimeSpan baseUtcOffset, string displayName, string standardDisplayName, string daylightDisplayName, TimeZoneInfo.AdjustmentRule[] adjustmentRules);

        //[Description("Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, a standard time name, a daylight saving time name, daylight saving time rules, and a value that indicates whether the returned object reflects daylight saving time information.")]
        //TODO: TimeZoneInfo CreateCustomTimeZone3(string id, TimeSpan baseUtcOffset, string displayName, string standardDisplayName, string daylightDisplayName, TimeZoneInfo.AdjustmentRule[] adjustmentRules, bool disableDaylightSavingTime);

        public TimeZoneInfo FindSystemTimeZoneById(string pId)
        {
            return TimeZoneInfo.FindSystemTimeZoneById(pId);
        }

        public TimeZoneInfo FromSerializedString(string source)
        {
            return TimeZoneInfo.FromSerializedString(source);
        }

        public ReadOnlyCollection GetSystemTimeZones()
        {
            return TimeZoneInfo.GetSystemTimeZones();
        }

    }
}
