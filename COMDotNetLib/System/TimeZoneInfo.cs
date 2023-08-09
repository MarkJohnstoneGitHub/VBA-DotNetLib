// https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1
// https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/TimeZoneInfo.cs

using System.Runtime.InteropServices;
using GSystem = global::System; // https://stackoverflow.com/questions/5681537/namespace-conflict-in-c-sharp
using System;
using System.ComponentModel;
using DotNetLib.System.Collections;

using System.Runtime.InteropServices.ComTypes;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEFLAGS = System.Runtime.InteropServices.ComTypes.TYPEFLAGS;
using DotNetLib.System;

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
    //TypeLibTypeAttribute(TypeLibTypeFlags.FPreDeclId)] //The type is predefined. The client application should automatically create a single instance of the object that has this attribute. 
    [TypeLibType(TypeLibTypeFlags.FPreDeclId | TypeLibTypeFlags.FCanCreate )] //The type is predefined. The client application should automatically create a single instance of the object that has this attribute. 
    [ProgId("DotNetLib.System.TimeZoneInfo")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITimeZoneInfo))]
    public class TimeZoneInfo : ITimeZoneInfo, ITimeZoneInfoSingleton
    {
        private GSystem.TimeZoneInfo timeZoneInfoObject;

        private static ReadOnlyCollection cachedSystemTimeZones = PopulateAllSystemTimeZones();
        private static ITimeZoneInfo cachedLocal = new TimeZoneInfo(GSystem.TimeZoneInfo.Local);
        private static ITimeZoneInfo cachedUtc  = new TimeZoneInfo(GSystem.TimeZoneInfo.Utc); 

        // Constructors

        public TimeZoneInfo()
        {
        }

        internal TimeZoneInfo(GSystem.TimeZoneInfo timeZoneInfoObject)
        {
            this.TimeZoneInfoObject = timeZoneInfoObject;
        }

        // Properties

        internal GSystem.TimeZoneInfo TimeZoneInfoObject
        {
            get { return this.timeZoneInfoObject; }
            set { this.timeZoneInfoObject = value; }  // set method for instance of the TimeZoneInfo class is immutable.
        }

        public ITimeSpan BaseUtcOffset
        { 
            get {return new TimeSpan(this.timeZoneInfoObject.BaseUtcOffset);}
        }
        public string DaylightName => this.timeZoneInfoObject.DaylightName;

        public string DisplayName => this.timeZoneInfoObject.DisplayName;

        public string Id => this.timeZoneInfoObject.Id;


        // TODO: Update to check if cachedLocal is null, if null create TimeZoneInfo, return cachedLocal ??
        public ITimeZoneInfo Locale
        {
            get { return cachedLocal; } 
            //get { return new TimeZoneInfo(GSystem.TimeZoneInfo.Local); }
        }

        public string StandardName => this.timeZoneInfoObject.StandardName;

        public bool SupportsDaylightSavingTime => this.timeZoneInfoObject.SupportsDaylightSavingTime;


        // TODO: Update to check if cachedUtc is null, if null create TimeZoneInfo, return cachedUtc
        public ITimeZoneInfo Utc
        {
            get { return cachedUtc; }
            //get { return new TimeZoneInfo(GSystem.TimeZoneInfo.Utc); }
        }

        // Methods

        public void ClearCachedData()
        {
            GSystem.TimeZoneInfo.ClearCachedData();

            // TODO: update to change to set chached items to null
            //refresh cached static data
            cachedSystemTimeZones = PopulateAllSystemTimeZones();
            cachedLocal = new TimeZoneInfo(GSystem.TimeZoneInfo.Local);
            cachedUtc = new TimeZoneInfo(GSystem.TimeZoneInfo.Utc);
        }

        public IDateTime ConvertTime(DateTime dateTime, TimeZoneInfo destinationTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTime(dateTime.DateTimeObject, destinationTimeZone.TimeZoneInfoObject));
        }

        public IDateTimeOffset ConvertTime2(DateTimeOffset dateTimeOffset, TimeZoneInfo destinationTimeZone)
        {
            return new DateTimeOffset(GSystem.TimeZoneInfo.ConvertTime(dateTimeOffset.DateTimeOffsetObject, destinationTimeZone.timeZoneInfoObject));
        }

        public IDateTime ConvertTime3(DateTime dateTime, TimeZoneInfo sourceTimeZone, TimeZoneInfo destinationTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTime(dateTime.DateTimeObject, sourceTimeZone.TimeZoneInfoObject ,destinationTimeZone.TimeZoneInfoObject));
        }

        public IDateTime ConvertTimeBySystemTimeZoneId(DateTime dateTime, string destinationTimeZoneId)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime.DateTimeObject, destinationTimeZoneId));
        }

        public IDateTimeOffset ConvertTimeBySystemTimeZoneId2(DateTimeOffset dateTimeOffset, string destinationTimeZoneId)
        {
            return new DateTimeOffset(GSystem.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTimeOffset.DateTimeOffsetObject, destinationTimeZoneId));
        }

        public IDateTime ConvertTimeBySystemTimeZoneId3(DateTime dateTime, string sourceTimeZoneId, string destinationTimeZoneId)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime.DateTimeObject, sourceTimeZoneId, destinationTimeZoneId));
        }

        public IDateTime ConvertTimeFromUtc(DateTime dateTime, TimeZoneInfo destinationTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeFromUtc(dateTime.DateTimeObject, destinationTimeZone.TimeZoneInfoObject));
        }

        public IDateTime ConvertTimeToUtc(DateTime dateTime)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeToUtc(dateTime.DateTimeObject));
        }
        public IDateTime ConvertTimeToUtc2(DateTime dateTime, TimeZoneInfo sourceTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeToUtc(dateTime.DateTimeObject, sourceTimeZone.TimeZoneInfoObject));
        }

        public ITimeZoneInfo CreateCustomTimeZone(string pId, TimeSpan pBaseUtcOffset, string pDisplayName, string standardDisplayName)
        {
            return new TimeZoneInfo(GSystem.TimeZoneInfo.CreateCustomTimeZone(pId, pBaseUtcOffset.TimeSpanObject, pDisplayName, standardDisplayName));
        }

        // TODO: TimeZoneInfo CreateCustomTimeZone2(string pId, TimeSpan pBaseUtcOffset, string pDisplayName, string standardDisplayName, string daylightDisplayName, TimeZoneInfo.AdjustmentRule[] adjustmentRules);
        // TODO: TimeZoneInfo CreateCustomTimeZone3(string pId, TimeSpan pBaseUtcOffset, string pDisplayName, string standardDisplayName, string daylightDisplayName, TimeZoneInfo.AdjustmentRule[] adjustmentRules, bool disableDaylightSavingTime);

        public bool Equals(TimeZoneInfo other)
        {
            return this.timeZoneInfoObject.Equals(other.TimeZoneInfoObject);
        }

        public bool Equals2(object obj)
        {
            if (!(obj is TimeZoneInfo tzi))
            {
                return false;
            }
            return Equals(tzi);
        }

        public ITimeZoneInfo FindSystemTimeZoneById(string pId)
        {
            return new TimeZoneInfo(GSystem.TimeZoneInfo.FindSystemTimeZoneById(pId));
        }

        public ITimeZoneInfo FromSerializedString(string source)
        {
            return new TimeZoneInfo(GSystem.TimeZoneInfo.FromSerializedString(source));
        }

        // TODO: public TimeZoneInfo.AdjustmentRule[] GetAdjustmentRules();

        public ITimeSpan[] GetAmbiguousTimeOffsets(DateTime dateTime)
        {

            GSystem.TimeSpan[] offsets = this.timeZoneInfoObject.GetAmbiguousTimeOffsets(dateTime.DateTimeObject);

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

        public ITimeSpan[] GetAmbiguousTimeOffsets2(DateTimeOffset dateTimeOffset)
        {
            GSystem.TimeSpan[] offsets = this.timeZoneInfoObject.GetAmbiguousTimeOffsets(dateTimeOffset.DateTimeOffsetObject);

            // Convert GSystem.TimeSpan[] offsets to DotNetLib.TimeSpan[]
            ITimeSpan[] timeSpans = new TimeSpan[offsets.Length];
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
            return this.timeZoneInfoObject.GetHashCode();
        }

        // As TimeZoneInfo.GetSystemTimeZones() returns a generic type ReadOnlyCollection<T> convert to non-generic type ReadOnlyCollectionBase
        // Note how Method ClearCache  and ROC<T> systemTimeZones is updated
        // Ideally only require to update the system time zone collection when the ReadOnlyCollection is updated
        // https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/TimeZoneInfo.cs,443c9b06db11142b

        // TODO: Update to check if cachedSystemTimeZones is null, if null cachedSystemTimeZones = PopulateAllSystemTimeZones, return cachedLocal
        public ReadOnlyCollection GetSystemTimeZones()
        {
            return cachedSystemTimeZones;
        }

        public ITimeSpan GetUtcOffset(DateTime dateTime)
        {
            return new TimeSpan(timeZoneInfoObject.GetUtcOffset(dateTime.DateTimeObject));
        }

        public ITimeSpan GetUtcOffset2(DateTimeOffset dateTimeOffset)
        {
            return new TimeSpan(timeZoneInfoObject.GetUtcOffset(dateTimeOffset.DateTimeOffsetObject));
        }

        public bool HasSameRules(TimeZoneInfo other)
        {
            return this.timeZoneInfoObject.HasSameRules(other.timeZoneInfoObject);
        }

        public bool IsAmbiguousTime(DateTime dateTime)
        {
            return this.timeZoneInfoObject.IsAmbiguousTime(dateTime.DateTimeObject);
        }

        public bool IsAmbiguousTime2(DateTimeOffset dateTimeOffset)
        {
            return this.timeZoneInfoObject.IsAmbiguousTime(dateTimeOffset.DateTimeOffsetObject);
        }

        public bool IsDaylightSavingTime(DateTime dateTime)
        {
            return this.timeZoneInfoObject.IsDaylightSavingTime(dateTime.DateTimeObject);
        }

        public bool IsDaylightSavingTime2(DateTimeOffset dateTimeOffset)
        {
            return this.timeZoneInfoObject.IsDaylightSavingTime(dateTimeOffset.DateTimeOffsetObject);
        }

        public bool IsInvalidTime(DateTime dateTime)
        {
            return this.timeZoneInfoObject.IsInvalidTime(dateTime.DateTimeObject);
        }

        public string ToSerializedString()
        {
            return this.timeZoneInfoObject.ToSerializedString();
        }

        public override string ToString()
        { 
            return this.timeZoneInfoObject.ToString();
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

