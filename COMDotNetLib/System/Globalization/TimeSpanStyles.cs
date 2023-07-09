using System;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    // https://learn.microsoft.com/en-us/dotnet/api/system.globalization.timespanstyles?view=netframework-4.8.1
    // https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/Globalization/TimeSpanStyles.cs

    [ComVisible(true)]
    [Guid("059973CB-4CB7-4747-A87D-3775AF981BB7")]

    [Flags]
    public enum TimeSpanStyles
    {
        None = 0x00000000,
        AssumeNegative = 0x00000001,
    }
}
