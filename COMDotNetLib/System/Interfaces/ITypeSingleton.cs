// https://learn.microsoft.com/en-us/dotnet/api/system.type?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("6EE512FF-1691-4A0A-B314-6DAC45F7F353")]
    [Description("Represents type declarations: class types, interface types, array types, value types, enumeration types, type parameters, generic type definitions, and open or closed constructed generic types.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITypeSingleton
    {
        [Description("Gets the Type with the specified name, performing a case-sensitive search.")]
        Type GetType(string typeName);

        [Description("Gets the Type with the specified name, performing a case-sensitive search and specifying whether to throw an exception if the type is not found.")]
        Type GetType(string typeName, bool throwOnError);

        [Description("Gets the Type with the specified name, specifying whether to throw an exception if the type is not found and whether to perform a case-sensitive search.")]
        Type GetType(string typeName, bool throwOnError, bool ignoreCase);

        [Description("Gets the underlying type code of the specified Type.")]
        TypeCode GetTypeCode(Type type);
    }
}
