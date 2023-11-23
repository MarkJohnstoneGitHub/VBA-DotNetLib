// https://learn.microsoft.com/en-us/dotnet/api/system.text.utf7encoding?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Description("Represents a UTF-7 encoding of Unicode characters.")]
    [Guid("B666D7DB-995A-4FBA-9684-EC8F0BAE6497")]
    [ProgId("DotNetLib.System.Text.UTF7EncodingSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IUTF7EncodingSingleton))]
    public class UTF7EncodingSingleton : IUTF7EncodingSingleton
    {
        public UTF7EncodingSingleton() { }

        //public UTF7Encoding Create()
        //{  
        //    return new UTF7Encoding(); 
        //}

        public UTF7Encoding Create(bool allowOptionals = false)
        { 
            return new UTF7Encoding(allowOptionals); 
        }

    }
}
