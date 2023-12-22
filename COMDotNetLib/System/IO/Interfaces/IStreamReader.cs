// https://learn.microsoft.com/en-us/dotnet/api/system.io.streamreader?view=netframework-4.8.1

using Encoding = DotNetLib.System.Text.Encoding;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("778E5959-AE19-4637-A8A9-B163E1436E1E")]
    [Description("Implements a TextReader that reads characters from a byte stream in a particular encoding.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]

    public interface IStreamReader
    {
        [Description("Returns the underlying stream.")]
        Stream BaseStream 
        { 
            get;
        }

        [Description("Gets the current character encoding that the current StreamReader object is using.")]
        Encoding CurrentEncoding 
        { 
            get;
        }

        [Description("Gets a value that indicates whether the current stream position is at the end of the stream.")]
        bool EndOfStream 
        { 
            get;
        }




    }
}
