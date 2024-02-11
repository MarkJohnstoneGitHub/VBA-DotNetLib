// https://learn.microsoft.com/en-us/dotnet/api/system.predicate-1?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("116FF3C7-C6D4-407E-8C42-1E5F233C5A8F")]
    [Description("Represents the method that defines a set of criteria and determines whether the specified object meets those criteria.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IPredicate
    {
        [Description("Represents the method that defines a set of criteria and determines whether the specified object meets those criteria.")]
        bool CallBack(object obj);

    }
}
