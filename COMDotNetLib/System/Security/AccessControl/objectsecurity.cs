using System.Runtime.InteropServices;

namespace DotNetLib.System.Security.AccessControl
{
    // https://learn.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.accesscontrolmodification?view=netframework-4.8.1

    [ComVisible(true)]
    public enum AccessControlModification
    {
        Add = 0,
        Set = 1,
        Reset = 2,
        Remove = 3,
        RemoveAll = 4,
        RemoveSpecific = 5,
    }


}
