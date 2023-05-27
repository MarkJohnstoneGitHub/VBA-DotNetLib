using System.Runtime.InteropServices;
using GSystem = global::System; // https://stackoverflow.com/questions/5681537/namespace-conflict-in-c-sharp
using System;
using System.ComponentModel;
using DotNetLib.System.Collections;

namespace DotNetLib.System
{

    // https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1
    // https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/TimeZoneInfo.cs

    [ComVisible(true)]
    [Description("Represents any time zone in the world.")]
    [Guid("A27D393F-5F4D-4F9B-9A5C-A72D980C802A")]
    [ProgId("DotNetLib.System.TimeZoneInfo")]
    [ClassInterface(ClassInterfaceType.None)]
    public class TimeZoneInfo : ITimeZoneInfo
    {
        private GSystem.TimeZoneInfo objTimeZoneInfo;

        private static ReadOnlyCollection cachedSystemTimeZones = PopulateAllSystemTimeZones();
        private static TimeZoneInfo cachedLocal = new TimeZoneInfo(GSystem.TimeZoneInfo.Local);
        private static TimeZoneInfo cachedUtc  = new TimeZoneInfo(GSystem.TimeZoneInfo.Utc); 

        public TimeZoneInfo()
        {
        }

        public TimeZoneInfo(GSystem.TimeZoneInfo objTimeZoneInfo)
        {
            this.objTimeZoneInfo = objTimeZoneInfo;
        }

        public GSystem.TimeZoneInfo timeZoneInfo
        {
            get { return this.objTimeZoneInfo; }
            set { objTimeZoneInfo = value; }  // set method
        }

        // Properties
        public TimeSpan BaseUtcOffset
        { 
            get {return new TimeSpan(this.objTimeZoneInfo.BaseUtcOffset);}
        }
        public string DaylightName => this.objTimeZoneInfo.DaylightName;

        public string DisplayName => this.objTimeZoneInfo.DisplayName;

        public string Id => this.objTimeZoneInfo.Id;

        public TimeZoneInfo Local
        {
            get { return new TimeZoneInfo(GSystem.TimeZoneInfo.Local); }
        }

        public string StandardName => this.objTimeZoneInfo.StandardName;

        public bool SupportsDaylightSavingTime => this.objTimeZoneInfo.SupportsDaylightSavingTime;
        public TimeZoneInfo Utc
        {
            get { return new TimeZoneInfo(GSystem.TimeZoneInfo.Utc); }
        }

        //Methods

        public void ClearCachedData()
        {
            GSystem.TimeZoneInfo.ClearCachedData();

            //refresh cached static data
            cachedSystemTimeZones = PopulateAllSystemTimeZones();
            cachedLocal = new TimeZoneInfo(GSystem.TimeZoneInfo.Local);
            cachedUtc = new TimeZoneInfo(GSystem.TimeZoneInfo.Utc);
        }

        public DateTime ConvertTime(DateTime dateTime, TimeZoneInfo destinationTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTime(dateTime.dateTime, destinationTimeZone.timeZoneInfo));
        }

        public DateTimeOffset ConvertTime2(DateTimeOffset dateTimeOffset, TimeZoneInfo destinationTimeZone)
        {
            return new DateTimeOffset(GSystem.TimeZoneInfo.ConvertTime(dateTimeOffset.dateTimeOffset, destinationTimeZone.objTimeZoneInfo));
        }

        public DateTime ConvertTime3(DateTime dateTime, TimeZoneInfo sourceTimeZone, TimeZoneInfo destinationTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTime(dateTime.dateTime, sourceTimeZone.timeZoneInfo ,destinationTimeZone.timeZoneInfo));
        }

        public DateTime ConvertTimeBySystemTimeZoneId(DateTime dateTime, string destinationTimeZoneId)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime.dateTime, destinationTimeZoneId));
        }

        public DateTimeOffset ConvertTimeBySystemTimeZoneId2(DateTimeOffset dateTimeOffset, string destinationTimeZoneId)
        {
            return new DateTimeOffset(GSystem.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTimeOffset.dateTimeOffset, destinationTimeZoneId));
        }

        public DateTime ConvertTimeBySystemTimeZoneId3(DateTime dateTime, string sourceTimeZoneId, string destinationTimeZoneId)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime.dateTime, sourceTimeZoneId, destinationTimeZoneId));
        }

        public DateTime ConvertTimeFromUtc(DateTime dateTime, TimeZoneInfo destinationTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeFromUtc(dateTime.dateTime, destinationTimeZone.timeZoneInfo));
        }

        public DateTime ConvertTimeToUtc(DateTime dateTime)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeToUtc(dateTime.dateTime));
        }
        public DateTime ConvertTimeToUtc2(DateTime dateTime, TimeZoneInfo sourceTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeToUtc(dateTime.dateTime, sourceTimeZone.timeZoneInfo));
        }

        public TimeZoneInfo CreateCustomTimeZone(string id, TimeSpan baseUtcOffset, string displayName, string standardDisplayName)
        {
            return new TimeZoneInfo(GSystem.TimeZoneInfo.CreateCustomTimeZone(id, baseUtcOffset.timeSpan, displayName, standardDisplayName));
        }

        //TimeZoneInfo CreateCustomTimeZone2(string id, TimeSpan baseUtcOffset, string displayName, string standardDisplayName, string daylightDisplayName, TimeZoneInfo.AdjustmentRule[] adjustmentRules);
        //TimeZoneInfo CreateCustomTimeZone3(string id, TimeSpan baseUtcOffset, string displayName, string standardDisplayName, string daylightDisplayName, TimeZoneInfo.AdjustmentRule[] adjustmentRules, bool disableDaylightSavingTime);


        public bool Equals(TimeZoneInfo other)
        {
            return this.objTimeZoneInfo.Equals(other.timeZoneInfo);
        }

        public bool Equals2(object obj)
        {
            return this.objTimeZoneInfo.Equals(obj);
        }

        public TimeZoneInfo FindSystemTimeZoneById(string id)
        {
            return new TimeZoneInfo(GSystem.TimeZoneInfo.FindSystemTimeZoneById(id));
        }

        public TimeZoneInfo FromSerializedString(string source)
        {
            return new TimeZoneInfo(GSystem.TimeZoneInfo.FromSerializedString(source));
        }

        //public TimeZoneInfo.AdjustmentRule[] GetAdjustmentRules();

        public TimeSpan[] GetAmbiguousTimeOffsets(DateTime dateTime)
        {

            GSystem.TimeSpan[] offsets = this.objTimeZoneInfo.GetAmbiguousTimeOffsets(dateTime.dateTime);

            // Convert GSystem.TimeSpan[] offsets to DotNetLib.TimeSpan[]
            TimeSpan[] timeSpans = new TimeSpan[offsets.Length];
            int i = 0;
            foreach (GSystem.TimeSpan offset in offsets)
            {
                timeSpans[i++].timeSpan = offset;
            }
            return timeSpans;
        }

        public TimeSpan[] GetAmbiguousTimeOffsets2(DateTimeOffset dateTimeOffset)
        {
            GSystem.TimeSpan[] offsets = this.objTimeZoneInfo.GetAmbiguousTimeOffsets(dateTimeOffset.dateTimeOffset);

            // Convert GSystem.TimeSpan[] offsets to DotNetLib.TimeSpan[]
            TimeSpan[] timeSpans = new TimeSpan[offsets.Length];
            int i = 0;
            foreach (GSystem.TimeSpan offset in offsets)
            {
                timeSpans[i++].timeSpan = offset;
            }
            return timeSpans;
        }

        public override int GetHashCode()
        { 
            return this.objTimeZoneInfo.GetHashCode();
        }


        // As TimeZoneInfo.GetSystemTimeZones() returns a generic type ReadOnlyCollection<T> convert to non-generic type ReadOnlyCollectionBase
        // Note how Method ClearCache  and ROC<T> systemTimeZones is updated
        // Ideally only require to update the system time zone collection when the ReadOnlyCollection is updated
        // https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/TimeZoneInfo.cs,443c9b06db11142b
        public ReadOnlyCollection GetSystemTimeZones()
        {
            return cachedSystemTimeZones;
        }

        public TimeSpan GetUtcOffset(DateTime dateTime)
        {
            return new TimeSpan(objTimeZoneInfo.GetUtcOffset(dateTime.dateTime));
        }

        public TimeSpan GetUtcOffset2(DateTimeOffset dateTimeOffset)
        {
            return new TimeSpan(objTimeZoneInfo.GetUtcOffset(dateTimeOffset.dateTimeOffset));
        }

        public bool HasSameRules(TimeZoneInfo other)
        {
            return this.objTimeZoneInfo.HasSameRules(other.objTimeZoneInfo);
        }

        public bool IsAmbiguousTime(DateTime dateTime)
        {
            return this.objTimeZoneInfo.IsAmbiguousTime(dateTime.dateTime);
        }

        public bool IsAmbiguousTime2(DateTimeOffset dateTimeOffset)
        {
            return this.objTimeZoneInfo.IsAmbiguousTime(dateTimeOffset.dateTimeOffset);
        }

        public bool IsDaylightSavingTime(DateTime dateTime)
        {
            return this.objTimeZoneInfo.IsDaylightSavingTime(dateTime.dateTime);
        }

        public bool IsDaylightSavingTime2(DateTimeOffset dateTimeOffset)
        {
            return this.objTimeZoneInfo.IsDaylightSavingTime(dateTimeOffset.dateTimeOffset);
        }

        public bool IsInvalidTime(DateTime dateTime)
        {
            return this.objTimeZoneInfo.IsInvalidTime(dateTime.dateTime);
        }

        public string ToSerializedString()
        {
            return this.objTimeZoneInfo.ToSerializedString();
        }

        public override string ToString()
        { 
            return this.objTimeZoneInfo.ToString();
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
