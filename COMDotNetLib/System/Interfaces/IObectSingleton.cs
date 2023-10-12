// https://learn.microsoft.com/en-us/dotnet/api/system.object?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("A201AB7E-B7FD-4C75-B186-C8A50BE45287")]
    [Description("Supports all classes in the .NET class hierarchy and provides low-level services to derived classes. This is the ultimate base class of all .NET classes; it is the root of the type hierarchy.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IObectSingleton
    {
        [Description("")]
        Object Create(object obj = null);

        [Description("Determines whether the specified object instances are considered equal.")]
        bool Equals(object objA, object objB);

        [Description("Determines whether the specified Object instances are the same instance.")]
        bool ReferenceEquals(object objA, object objB);
    }
}
