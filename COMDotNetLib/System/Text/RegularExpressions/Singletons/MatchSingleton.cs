// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.match?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("326EB36E-CEBB-4DF4-A9EC-D85FDE7D09E7")]
    [ProgId("DotNetLib.System.Text.RegularExpressions.MatchSingleton")]
    [Description("Represents the results from a single regular expression match.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMatchSingleton))]
    public class MatchSingleton : IMatchSingleton
    {
        public MatchSingleton() { }

        // Properties
        public Match EmptyMatch => Match.Empty;

        // Methods
        public  Match Synchronized(Match inner)
        {  
            return Match.Synchronized(inner); 
        }
    }
}
