// https://learn.microsoft.com/en-us/dotnet/api/system.typecode?view=netframework-4.8.1

using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Description("Type Code helper to convert a value of this instance to its equivalent string representation.")]
    [Guid("C28B90C6-FBDA-4942-94F7-CCEED784B9F5")]
    [ProgId("DotNetLib.System.TypeCodeHelperSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITypeCodeHelperSingleton))]
    public class TypeCodeHelperSingleton : ITypeCodeHelperSingleton
    {
        public TypeCodeHelperSingleton() { }

        public string ToString(GSystem.TypeCode typecode)
        {
            return typecode.ToString();
        }

        public string ToString2(GSystem.TypeCode typecode, string format)
        {
            return typecode.ToString(format);
        }

    }
}
