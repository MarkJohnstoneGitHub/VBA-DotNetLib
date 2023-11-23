// https://learn.microsoft.com/en-us/dotnet/api/system.text.asciiencoding?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Guid("B8F6ED97-56EA-4AD5-A38C-D9DC6919DB4E")]
    [Description("Represents an ASCII character encoding of Unicode characters.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IASCIIEncodingSingleton
    {
        [Description("Initializes a new instance of the ASCIIEncoding class.")]
        ASCIIEncoding Create();
    }
}
