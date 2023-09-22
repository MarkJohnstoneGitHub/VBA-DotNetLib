// https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

using DotNetLib.Extensions;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace DotNetLib.System.Globalization
{

    [ComVisible(true)]
    [Description("")]
    [Guid("2BB0ED15-8B6E-4D70-9DA2-A1C1BA9F8EC3")]
    [ProgId("DotNetLib.System.StringSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStringSingleton))]

    public class StringSingleton : IStringSingleton
    {
        public StringSingleton() { }

        public string Format(string pFormat, [In] ref object[] args)
        {
            return string.Format(pFormat, args.Unwrap());
        }
        public string Format2(IFormatProvider provider, string pFormat, [In] ref object[] args)
        {
            return string.Format(provider.Unwrap() , pFormat, args.Unwrap());
        }

        //public string Format(string format, object arg0)
        //{  
        //    return string.Format(format, arg0.Unwrap()); 
        //}

        //public string Format(string format, object arg0, object arg1)
        //{
        //    return string.Format(format, arg0.Unwrap(), arg1.Unwrap());
        //}

        //public string Format(string format, object arg0, object arg1, object arg2)
        //{
        //    return string.Format(format, arg0.Unwrap(), arg1.Unwrap(), arg2.Unwrap());
        //}

        //public string Format(string format, object arg0, object arg1, object arg2, object arg3)
        //{
        //    return string.Format(format, arg0.Unwrap(), arg1.Unwrap(), arg2.Unwrap(), arg3.Unwrap());
        //}

        //public string Format(IFormatProvider provider, string format, object arg0)
        //{
        //    return string.Format(provider, format, arg0.Unwrap());
        //}

        //public static string Format(IFormatProvider provider, string format, object arg0, object arg1)
        //{
        //    return string.Format(provider, format, arg0.Unwrap(), arg1.Unwrap());
        //}

        //public string Format(IFormatProvider provider, string format, object arg0, object arg1, object arg2)
        //{
        //    return string.Format(provider, format, arg0.Unwrap(), arg1.Unwrap(), arg2.Unwrap());
        //}

    }
}
