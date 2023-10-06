// https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1
// https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/TimeZoneInfo.cs

using System.Runtime.InteropServices;
using GSystem = global::System;
using System;
using System.ComponentModel;
using DotNetLib.System.Collections;

namespace DotNetLib.System
{
    // TODO : Explict Interface Implementations
    // TODO: ClearCachedData() remove cachedLocal, cachedUtc, PopulateAllSystemTimeZones
    // TODO: For Local property check if cachedLocal is null, if null create TimeZoneInfo object and return cachedLocal
    // TODO: For Utc property check if cachedUtc is null, if null create TimeZoneInfo and return TimeZoneInfo
    // TODO: For GetSystemTimeZones check if cachedSystemTimeZones is null, if null PopulateAllSystemTimeZones
    // TODO: Require to examine how to implement the cachedData see ClearCachedData in .net source code and how it is handled.


    // An instance of the TimeZoneInfo class is immutable.Once an object has been instantiated, its values cannot be modified.

    // You cannot instantiate a TimeZoneInfo object using the new keyword.Instead, you must call one of the static members of the TimeZoneInfo class shown in the following table.

    // Static member name Description
    // CreateCustomTimeZone method Creates a custom time zone from application-supplied data.
    // FindSystemTimeZoneById method Instantiates a time zone based on its identifier.
    // FromSerializedString method Deserializes a string value to re-create a previously serialized TimeZoneInfo object.
    // GetSystemTimeZones method   Returns an enumerable ReadOnlyCollection<T> of TimeZoneInfo objects that represents all time zones that are available on the local system.
    // Local property  Instantiates a TimeZoneInfo object that represents the local time zone.
    // Utc property    Instantiates a TimeZoneInfo object that represents the UTC zone.

    [ComVisible(true)]
    [Description("Represents any time zone in the world.")]
    [Guid("A27D393F-5F4D-4F9B-9A5C-A72D980C802A")]
    [ProgId("DotNetLib.System.TimeZoneInfo")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITimeZoneInfo))]
    public class TimeZoneInfo : ITimeZoneInfo
    {
        private GSystem.TimeZoneInfo _timeZoneInfo;

        private static ReadOnlyCollection cachedSystemTimeZones = PopulateAllSystemTimeZones();
        private static TimeZoneInfo cachedLocal = new TimeZoneInfo(GSystem.TimeZoneInfo.Local);
        private static TimeZoneInfo cachedUtc  = new TimeZoneInfo(GSystem.TimeZoneInfo.Utc); 

        // Constructors

        //public TimeZoneInfo()
        //{
        //}

        internal TimeZoneInfo(GSystem.TimeZoneInfo timeZoneInfo)
        {
            WrappedTimeZoneInfo = timeZoneInfo;
        }

        // Properties

        internal GSystem.TimeZoneInfo WrappedTimeZoneInfo
        {
            get => _timeZoneInfo;
            set { _timeZoneInfo = value; }  // set method for instance of the TimeZoneInfo class is immutable.
        }

        public TimeSpan BaseUtcOffset => new TimeSpan(_timeZoneInfo.BaseUtcOffset);  //TODO: Check

        public string DaylightName => _timeZoneInfo.DaylightName;

        public string DisplayName => _timeZoneInfo.DisplayName;

        public string Id => _timeZoneInfo.Id;


        // TODO: SetItem to check if cachedLocal is null, if null create TimeZoneInfo, return cachedLocal ??
        public static TimeZoneInfo Local => cachedLocal; 

        public string StandardName => _timeZoneInfo.StandardName;

        public bool SupportsDaylightSavingTime => _timeZoneInfo.SupportsDaylightSavingTime;

        // TODO: SetItem to check if cachedUtc is null, if null create TimeZoneInfo, return cachedUtc
        public static TimeZoneInfo Utc => cachedUtc;


        // Methods

        public static void ClearCachedData()
        {
            GSystem.TimeZoneInfo.ClearCachedData();

            // TODO: update to change to set chached items to null
            //refresh cached static data
            cachedSystemTimeZones = PopulateAllSystemTimeZones();
            cachedLocal = new TimeZoneInfo(GSystem.TimeZoneInfo.Local);
            cachedUtc = new TimeZoneInfo(GSystem.TimeZoneInfo.Utc);
        }

        public static DateTime ConvertTime(DateTime dateTime, TimeZoneInfo destinationTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTime(dateTime.WrappedDateTime, destinationTimeZone.WrappedTimeZoneInfo));
        }

        public static DateTimeOffset ConvertTime(DateTimeOffset dateTimeOffset, TimeZoneInfo destinationTimeZone)
        {
            return new DateTimeOffset(GSystem.TimeZoneInfo.ConvertTime(dateTimeOffset.WrappedDateTimeOffset, destinationTimeZone._timeZoneInfo));
        }

        public static DateTime ConvertTime(DateTime dateTime, TimeZoneInfo sourceTimeZone, TimeZoneInfo destinationTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTime(dateTime.WrappedDateTime, sourceTimeZone.WrappedTimeZoneInfo ,destinationTimeZone.WrappedTimeZoneInfo));
        }

        public static DateTime ConvertTimeBySystemTimeZoneId(DateTime dateTime, string destinationTimeZoneId)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime.WrappedDateTime, destinationTimeZoneId));
        }

        public static DateTimeOffset ConvertTimeBySystemTimeZoneId(DateTimeOffset dateTimeOffset, string destinationTimeZoneId)
        {
            return new DateTimeOffset(GSystem.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTimeOffset.WrappedDateTimeOffset, destinationTimeZoneId));
        }

        public static DateTime ConvertTimeBySystemTimeZoneId(DateTime dateTime, string sourceTimeZoneId, string destinationTimeZoneId)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime.WrappedDateTime, sourceTimeZoneId, destinationTimeZoneId));
        }

        public static DateTime ConvertTimeFromUtc(DateTime dateTime, TimeZoneInfo destinationTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeFromUtc(dateTime.WrappedDateTime, destinationTimeZone.WrappedTimeZoneInfo));
        }

        public static DateTime ConvertTimeToUtc(DateTime dateTime)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeToUtc(dateTime.WrappedDateTime));
        }
        public static DateTime ConvertTimeToUtc(DateTime dateTime, TimeZoneInfo sourceTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeToUtc(dateTime.WrappedDateTime, sourceTimeZone.WrappedTimeZoneInfo));
        }

        public static TimeZoneInfo CreateCustomTimeZone(string pId, TimeSpan pBaseUtcOffset, string pDisplayName, string standardDisplayName)
        {
            return new TimeZoneInfo(GSystem.TimeZoneInfo.CreateCustomTimeZone(pId, pBaseUtcOffset.WrappedTimeSpan, pDisplayName, standardDisplayName));
        }

        // TODO: TimeZoneInfo CreateCustomTimeZone2(string pId, TimeSpan pBaseUtcOffset, string pDisplayName, string standardDisplayName, string daylightDisplayName, TimeZoneInfo.AdjustmentRule[] adjustmentRules);
        // TODO: TimeZoneInfo CreateCustomTimeZone3(string pId, TimeSpan pBaseUtcOffset, string pDisplayName, string standardDisplayName, string daylightDisplayName, TimeZoneInfo.AdjustmentRule[] adjustmentRules, bool disableDaylightSavingTime);

        public bool Equals(TimeZoneInfo other)
        {
            return _timeZoneInfo.Equals(other.WrappedTimeZoneInfo);
        }

        public bool Equals2(object obj)
        {
            if (!(obj is TimeZoneInfo tzi))
            {
                return false;
            }
            return Equals(tzi);
        }

        public static TimeZoneInfo FindSystemTimeZoneById(string pId)
        {
            return new TimeZoneInfo(GSystem.TimeZoneInfo.FindSystemTimeZoneById(pId));
        }

        public static TimeZoneInfo FromSerializedString(string source)
        {
            return new TimeZoneInfo(GSystem.TimeZoneInfo.FromSerializedString(source));
        }

        // TODO: public TimeZoneInfo.AdjustmentRule[] GetAdjustmentRules();

        public TimeSpan[] GetAmbiguousTimeOffsets(DateTime dateTime)
        {

            GSystem.TimeSpan[] offsets = _timeZoneInfo.GetAmbiguousTimeOffsets(dateTime.WrappedDateTime);

            // Convert GSystem.TimeSpan[] offsets to DotNetLib.TimeSpan[]
            TimeSpan[] timeSpans = new TimeSpan[offsets.Length];
            int i = 0;
            foreach (GSystem.TimeSpan offset in offsets)
            {
                timeSpans[i] = new TimeSpan(offset);
                i++;
            }
            return timeSpans;
        }

        public TimeSpan[] GetAmbiguousTimeOffsets2(DateTimeOffset dateTimeOffset)
        {
            GSystem.TimeSpan[] offsets = _timeZoneInfo.GetAmbiguousTimeOffsets(dateTimeOffset.WrappedDateTimeOffset);

            // Convert GSystem.TimeSpan[] offsets to DotNetLib.TimeSpan[]
            TimeSpan[] timeSpans = new TimeSpan[offsets.Length];
            int i = 0;
            foreach (GSystem.TimeSpan offset in offsets)
            {
                timeSpans[i] = new TimeSpan(offset);
                i++;
            }
            return timeSpans;
        }

        public override int GetHashCode()
        { 
            return _timeZoneInfo.GetHashCode();
        }

        // As TimeZoneInfo.GetSystemTimeZones() returns a generic type ReadOnlyCollection<T> convert to non-generic type ReadOnlyCollectionBase
        // Note how Method ClearCache  and ROC<T> systemTimeZones is updated
        // Ideally only require to update the system time zone collection when the ReadOnlyCollection is updated
        // https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/TimeZoneInfo.cs,443c9b06db11142b

        // TODO: SetItem to check if cachedSystemTimeZones is null, if null cachedSystemTimeZones = PopulateAllSystemTimeZones, return cachedLocal
        public static ReadOnlyCollection GetSystemTimeZones()
        {
            return cachedSystemTimeZones;
        }

        public TimeSpan GetUtcOffset(DateTime dateTime)
        {
            return new TimeSpan(_timeZoneInfo.GetUtcOffset(dateTime.WrappedDateTime));
        }

        public TimeSpan GetUtcOffset2(DateTimeOffset dateTimeOffset)
        {
            return new TimeSpan(_timeZoneInfo.GetUtcOffset(dateTimeOffset.WrappedDateTimeOffset));
        }

        public bool HasSameRules(TimeZoneInfo other)
        {
            return _timeZoneInfo.HasSameRules(other._timeZoneInfo);
        }

        public bool IsAmbiguousTime(DateTime dateTime)
        {
            return _timeZoneInfo.IsAmbiguousTime(dateTime.WrappedDateTime);
        }

        public bool IsAmbiguousTime2(DateTimeOffset dateTimeOffset)
        {
            return _timeZoneInfo.IsAmbiguousTime(dateTimeOffset.WrappedDateTimeOffset);
        }

        public bool IsDaylightSavingTime(DateTime dateTime)
        {
            return _timeZoneInfo.IsDaylightSavingTime(dateTime.WrappedDateTime);
        }

        public bool IsDaylightSavingTime2(DateTimeOffset dateTimeOffset)
        {
            return _timeZoneInfo.IsDaylightSavingTime(dateTimeOffset.WrappedDateTimeOffset);
        }

        public bool IsInvalidTime(DateTime dateTime)
        {
            return _timeZoneInfo.IsInvalidTime(dateTime.WrappedDateTime);
        }

        public string ToSerializedString()
        {
            return _timeZoneInfo.ToSerializedString();
        }

        public override string ToString()
        { 
            return _timeZoneInfo.ToString();
        }

        //Populate a ReadOnlyCollection of system time zones from ReadOnlyCollection<TimeZoneInfo> from TimeZoneInfo.GetSystemTimeZones()
        //Converts generic type ReadOnlyCollection<GSystem.TimeZoneInfo> to non-generic type ReadOnlyCollection
        private static ReadOnlyCollection PopulateAllSystemTimeZones()
        {
            GSystem.Collections.ObjectModel.ReadOnlyCollection<GSystem.TimeZoneInfo> systemTimeZones = GSystem.TimeZoneInfo.GetSystemTimeZones();
            GSystem.Collections.ArrayList timeZoneInfos = new GSystem.Collections.ArrayList(systemTimeZones.Count);
            foreach (GSystem.TimeZoneInfo systemTimeZone in systemTimeZones)
            {
                timeZoneInfos.Add(new TimeZoneInfo(systemTimeZone));
            }
            ReadOnlyCollection colSystemTimeZones = new ReadOnlyCollection(timeZoneInfos);
            return colSystemTimeZones;
        }
    }
}

