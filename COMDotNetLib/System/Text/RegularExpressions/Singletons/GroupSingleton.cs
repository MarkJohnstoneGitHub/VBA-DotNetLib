// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.group?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("814339DE-9D5D-44EB-9EF0-AA6488573E16")]
    [ProgId("DotNetLib.System.Text.RegularExpressions.GroupSingleton")]
    [Description("Represents the results from a single capturing group.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IGroupSingleton))]
    public class GroupSingleton : IGroupSingleton
    {
        public Group Synchronized(Group inner)
        {
            return Group.Synchronized(inner);
        }
    }
}
