using System.Runtime.InteropServices;

namespace DotNetLib.System.Security.AccessControl
{
    // https://learn.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.accesscontroltype?view=netframework-4.8.1

    [ComVisible(true)]
    public enum AccessControlType
    {
        Allow = 0,
        Deny = 1,
    }


}
