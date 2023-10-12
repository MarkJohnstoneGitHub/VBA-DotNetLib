// https://learn.microsoft.com/en-us/dotnet/api/system.type?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("A86490BC-5CA3-414B-B501-1A142B7A6EA6")]
    [ProgId("DotNetLib.System.TypeSingleton")]
    [Description("Represents type declarations: class types, interface types, array types, value types, enumeration types, type parameters, generic type definitions, and open or closed constructed generic types.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITypeSingleton))]
    public class TypeSingleton : ITypeSingleton
    {
        public TypeSingleton() { }

        public Type GetType(string typeName)
        {
            return Type.GetType(typeName);
        }

        public Type GetType(string typeName, bool throwOnError)
        {
            return Type.GetType(typeName, throwOnError);
        }

        public Type GetType(string typeName, bool throwOnError, bool ignoreCase)
        {
            return Type.GetType(typeName, throwOnError, ignoreCase);
        }

        public TypeCode GetTypeCode(Type type)
        {
            return Type.GetTypeCode(type);
        }
    }
}
