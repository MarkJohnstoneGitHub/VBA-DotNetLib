//  https://learn.microsoft.com/en-us/dotnet/api/system.diagnostics.stopwatch?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Diagnostics
{
    [ComVisible(true)]
    [Guid("7F3B631B-3FB1-4268-BDB1-4CF92FC0C078")]
    [Description("Provides a set of methods and properties that you can use to accurately measure elapsed time.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStopwatch
    {
        TimeSpan Elapsed 
        {
            [Description("Gets the total elapsed time measured by the current instance.")]
            get;
        }

        long ElapsedMilliseconds 
        {
            [Description("Gets the total elapsed time measured by the current instance, in milliseconds.")]
            get;
        }

        long ElapsedTicks 
        {
            [Description("Gets the total elapsed time measured by the current instance, in timer ticks.")]
            get;
        }

        bool IsRunning 
        {
            [Description("Gets a value indicating whether the Stopwatch timer is running.")]
            get;
        }

        [Description("Stops time interval measurement and resets the elapsed time to zero.")]
        void Reset();

        [Description("Stops time interval measurement, resets the elapsed time to zero, and starts measuring elapsed time.")]
        void Restart();

        [Description("Starts, or resumes, measuring elapsed time for an interval.")]
        void Start();

        [Description("Stops measuring elapsed time for an interval.")]
        void Stop();

    }
}
