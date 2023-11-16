using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNetLib.Extensions
{
    public static class StringExtension
    {

        // https://www.nuget.org/packages/System.Runtime.CompilerServices.Unsafe/
        // https://stackoverflow.com/a/274207/10759363
        // https://stackoverflow.com/questions/274158/c-sharp-ushort-to-string-conversion-is-this-possible
        // C# ushort[] to string conversion; is this possible?

        //public static string ConvertToString(this ushort[] uSpan)
        //{
        //    byte[] bytes = new byte[sizeof(ushort) * uSpan.Length];

        //    for (int i = 0; i < uSpan.Length; i++)
        //    {
        //        Unsafe.As<byte, ushort>(ref bytes[i * 2]) = uSpan[i];
        //    }

        //    return Encoding.Unicode.GetString(bytes);
        //}

    }
}
