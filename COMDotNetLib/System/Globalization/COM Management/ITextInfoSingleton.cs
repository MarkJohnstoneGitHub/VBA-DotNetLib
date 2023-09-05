// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.textinfo?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("DF106974-C116-41C9-B766-3D49D592C337")]
    [Description("TextInfo factory methods and static members.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITextInfoSingleton
    {
        [Description("Returns a read-only version of the specified TextInfo object.")]
        TextInfo ReadOnly(TextInfo textInfo);
    }
}
