// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.group?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("5FCE41A5-67FC-4E5F-A52D-0B24C42D53F8")]
    [Description("Represents the results from a single capturing group.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IGroupSingleton
    {
        [Description("Returns a Group object equivalent to the one supplied that is safe to share between multiple threads.")]
        Group Synchronized(Group inner);
    }
}
