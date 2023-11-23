// https://learn.microsoft.com/en-us/dotnet/api/system.text.asciiencoding?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Description("Represents an ASCII character encoding of Unicode characters.")]
    [Guid("35DE9C35-E7B9-430C-8A3B-BA71CF7383D0")]
    [ProgId("DotNetLib.System.Text.ASCIIEncodingSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IASCIIEncodingSingleton))]
    public class ASCIIEncodingSingleton : IASCIIEncodingSingleton
    {
        public ASCIIEncodingSingleton() { }

        public ASCIIEncoding Create()
        { 
            return new ASCIIEncoding(); 
        }

    }
}
