// https://learn.microsoft.com/en-us/dotnet/api/system.text.unicodeencoding?view=netframework-4.8.1

using GText = global::System.Text;
using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Description("Represents a UTF-16 encoding of Unicode characters.")]
    [Guid("D58AB57E-BD99-4BC8-8EF1-7C904922EA84")]
    [ProgId("DotNetLib.System.Text.UnicodeEncodingSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IUnicodeEncodingSingleton))]
    public class UnicodeEncodingSingleton : IUnicodeEncodingSingleton
    {
        public UnicodeEncodingSingleton() { }

        // Fields
        public int CharSize => GText.UnicodeEncoding.CharSize;

        // Factory Methods
        //public UnicodeEncoding Create()
        //{
        //    return new UnicodeEncoding();
        //}

        //public UnicodeEncoding Create(bool bigEndian, bool byteOrderMark)
        //{
        //    return new UnicodeEncoding(bigEndian, byteOrderMark);
        //}

        public UnicodeEncoding Create(bool bigEndian = false, bool byteOrderMark = true, bool throwOnInvalidBytes = false)
        {
            return new UnicodeEncoding(bigEndian, byteOrderMark, throwOnInvalidBytes);
        }



    }
}
