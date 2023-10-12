// https://learn.microsoft.com/en-us/dotnet/api/system.object?view=netframework-4.8.1

using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("C4A457F6-862D-432A-A03D-DA5D29B125DD")]
    [Description("Supports all classes in the .NET class hierarchy and provides low-level services to derived classes. This is the ultimate base class of all .NET classes; it is the root of the type hierarchy.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IObect 
    {
        [Description("Determines whether the specified object is equal to the current object.")]
        bool Equals(object obj);

        [Description("Serves as the default hash function.")]
        int GetHashCode();

        [Description("Gets the Type of the current instance.")]
        Type GetType();

        [Description("Returns a string that represents the current object.")]
        string ToString();

    }
}
