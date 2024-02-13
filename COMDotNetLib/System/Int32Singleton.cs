// https://learn.microsoft.com/en-us/dotnet/api/system.int32?view=netframework-4.8.1

using GSystem = global::System;
using GGlobalization = global::System.Globalization;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("76DDE809-70BF-4975-8FE2-3A537B41E5BA")]
    [ProgId("DotNetLib.System.Int32Singleton")]
    [Description("Represents a 32-bit signed integer.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IInt32Singleton))]
    public class Int32Singleton : IInt32Singleton
    {

        public Int32Singleton() { }

        //Fields

        public int MaxValue => GSystem.Int32.MaxValue;

        public int MinValue => GSystem.Int32.MinValue;


        public bool TryParse(string s, out int result)
        {
            return GSystem.Int32.TryParse(s, out result);
        }

        public bool TryParse(string s, GGlobalization.NumberStyles style, GSystem.IFormatProvider provider, out int result)
        {
            return GSystem.Int32.TryParse(s, style, provider, out result);
        }

    }
}
