//  https://learn.microsoft.com/en-us/dotnet/api/system.diagnostics.stopwatch?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Diagnostics
{
    [ComVisible(true)]
    [Guid("439985F1-2EC5-40B1-8566-2F7050B262C7")]
    [Description("Provides a set of methods and properties that you can use to accurately measure elapsed time.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStopwatchSingleton
    {
        // Factory Methods
        [Description("Initializes a new instance of the Stopwatch class.")]
        Stopwatch Create();

        // Fields
        long Frequency
        {
            [Description("Gets the frequency of the timer as the number of ticks per second. This field is read-only.")]
            get;
        }

        bool IsHighResolution
        {
            [Description("Indicates whether the timer is based on a high-resolution performance counter. This field is read-only.")]
            get;
        }

        // Methods
        [Description("Gets the current number of ticks in the timer mechanism.")]
        long GetTimestamp();

        [Description("Initializes a new Stopwatch instance, sets the elapsed time property to zero, and starts measuring elapsed time.")]
        Stopwatch StartNew();

    }
}
