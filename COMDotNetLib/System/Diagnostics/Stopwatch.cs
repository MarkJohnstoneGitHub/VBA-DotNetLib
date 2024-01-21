//  https://learn.microsoft.com/en-us/dotnet/api/system.diagnostics.stopwatch?view=netframework-4.8.1

using GDiagnostics = global::System.Diagnostics;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Diagnostics
{
    [ComVisible(true)]
    [Guid("23998C73-200E-410C-94D7-2C311A7209F0")]
    [ProgId("DotNetLib.System.Diagnostics.Stopwatch")]
    [Description("Provides a set of methods and properties that you can use to accurately measure elapsed time.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStopwatch))]

    public class Stopwatch : IStopwatch
    {
        private GDiagnostics.Stopwatch _stopwatch;

        public Stopwatch() 
        {
            _stopwatch = new GDiagnostics.Stopwatch();
        }

        internal Stopwatch(GDiagnostics.Stopwatch stopwatch) 
        {
            _stopwatch = stopwatch;
        }

        // Properties
        public TimeSpan Elapsed => new TimeSpan(_stopwatch.Elapsed);

        public long ElapsedMilliseconds => _stopwatch.ElapsedMilliseconds;

        public long ElapsedTicks => _stopwatch.ElapsedTicks;

        public bool IsRunning => _stopwatch.IsRunning;

        public void Reset()
        {
            _stopwatch.Reset();
        }

        public void Restart()
        { 
            _stopwatch.Restart(); 
        }

        public void Start()
        { 
            _stopwatch.Start();
        }
        public void Stop() 
        { 
            _stopwatch.Stop();
        }

    }
}
