// https://learn.microsoft.com/en-us/dotnet/api/system.predicate-1?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("FCAAD9F1-A9B0-4501-9BCC-E1F7A904BFB2")]
    [Description("Represents the method that defines a set of criteria and determines whether the specified object meets those criteria.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IPredicateSingleton
    {
        [Description("Assigns the predicate callback method to the Predicate delegate.")]
        Predicate Create(IPredicate predicate);
    }
}
