using System.Runtime.InteropServices;
using GSystem = global::System; // https://stackoverflow.com/questions/5681537/namespace-conflict-in-c-sharp
using DotNetLib.System;
using System.Runtime.InteropServices.WindowsRuntime;
using System;
using System.ComponentModel;

namespace DotNetLib.System
{

    // https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/TimeZoneInfo.cs

    [ComVisible(true)]
    [Description("Represents any time zone in the world.")]
    [Guid("A27D393F-5F4D-4F9B-9A5C-A72D980C802A")]
    [ProgId("DotNetLib.System.TimeZoneInfo")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class TimeZoneInfo
    {
        private GSystem.TimeZoneInfo objTimeZoneInfo;


        public TimeZoneInfo(GSystem.TimeZoneInfo objTimeZoneInfo)
        {
            this.objTimeZoneInfo = objTimeZoneInfo;
        }

        public GSystem.TimeZoneInfo timeZoneInfo
        {
            get { return this.objTimeZoneInfo; }
            set { objTimeZoneInfo = value; }  // set method
        }
        public TimeSpan BaseUtcOffset { get; }
        public string DaylightName { get; }
        public string DisplayName { get; }
        public string Id { get; }
        public static TimeZoneInfo Local { get; }
        public string StandardName { get; }
        public bool SupportsDaylightSavingTime { get; }
        public static TimeZoneInfo Utc { get; }


        //Methods
        public DateTime ConvertTime(DateTime dateTime, TimeZoneInfo destinationTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTime(dateTime.dateTime, destinationTimeZone.timeZoneInfo));
        }

        public DateTime ConvertTime2(DateTime dateTime, TimeZoneInfo sourceTimeZone, TimeZoneInfo destinationTimeZone)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTime(dateTime.dateTime, sourceTimeZone.timeZoneInfo ,destinationTimeZone.timeZoneInfo));
        }

        public DateTime ConvertTimeBySystemTimeZoneId(DateTime dateTime, string destinationTimeZoneId)
        {
            return new DateTime(GSystem.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dateTime.dateTime, destinationTimeZoneId));
        }

        public DateTime ConvertTimeBySystemTimeZoneId2(DateTime dateTime, string sourceTimeZoneId, string destinationTimeZoneId)
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

        public bool Equals(TimeZoneInfo other)
        {
            return this.objTimeZoneInfo.Equals(other.timeZoneInfo);
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

        //public TimeSpan[] GetAmbiguousTimeOffsets(DateTime dateTime)
        //{
        //    return new TimeSpan[0]; 
        //}

        public override int GetHashCode()
        { 
            return this.objTimeZoneInfo.GetHashCode();
        }

        public TimeSpan GetUtcOffset(DateTime dateTime)
        {
            return new TimeSpan(this.timeZoneInfo.GetUtcOffset(dateTime.dateTime));
        }



    }
}
