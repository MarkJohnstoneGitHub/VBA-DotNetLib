//  https://learn.microsoft.com/en-us/dotnet/api/system.diagnostics.stopwatch?view=netframework-4.8.1

using GDiagnostics = global::System.Diagnostics;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Diagnostics
{
    [ComVisible(true)]
    [Guid("98AE3EE0-CD12-4D4B-8F14-54F7624AD8D8")]
    [ProgId("DotNetLib.System.Diagnostics.StopwatchSingleton")]
    [Description("Provides a set of methods and properties that you can use to accurately measure elapsed time.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStopwatchSingleton))]

    public class StopwatchSingleton : IStopwatchSingleton
    {
        public StopwatchSingleton() { }

        // Factory Methods
        
        public Stopwatch Create()
        { 
            return new Stopwatch(); 
        }

        // Fields
        
        public long Frequency => GDiagnostics.Stopwatch.Frequency;

        public bool IsHighResolution => GDiagnostics.Stopwatch.IsHighResolution;

        // Methods

        public long GetTimestamp()
        {  
            return GDiagnostics.Stopwatch.GetTimestamp();
        }

        public Stopwatch StartNew()
        {
            return new Stopwatch(GDiagnostics.Stopwatch.StartNew());
        }
        

    }
}
