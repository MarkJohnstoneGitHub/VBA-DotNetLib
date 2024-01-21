// https://learn.microsoft.com/en-us/dotnet/api/system.text.utf8encoding?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Description("Represents a UTF-8 encoding of Unicode characters.")]
    [Guid("31429B1B-EA88-4796-BB69-DFD9E66DA8EB")]
    [ProgId("DotNetLib.System.Text.UTF8EncodingSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IUTF8EncodingSingleton))]
    public class UTF8EncodingSingleton : IUTF8EncodingSingleton
    {

        //public UTF8Encoding Create2()
        //{
        //    return new UTF8Encoding();
        //}

        //public UTF8Encoding Create2(bool encoderShouldEmitUTF8Identifier)
        //{
        //    return new UTF8Encoding(encoderShouldEmitUTF8Identifier);
        //}

        public UTF8Encoding Create(bool encoderShouldEmitUTF8Identifier = false, bool throwOnInvalidBytes = false)
        {
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier, throwOnInvalidBytes);
        }

    }
}
