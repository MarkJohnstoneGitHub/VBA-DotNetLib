// // https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.match?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("593AAA6F-72CA-4577-9E22-3A29187E7D7F")]
    [Description("Represents the results from a single regular expression match.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMatchSingleton
    {
        //Properties
        Match Empty
        {
            [Description("Gets the empty group. All failed matches return this empty match.")]
            get;
        }

        // Methods
        [Description("Returns a Match instance equivalent to the one supplied that is suitable to share between multiple threads.")]
        Match Synchronized(Match inner);

    }
}
