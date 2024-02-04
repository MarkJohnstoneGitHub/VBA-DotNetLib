// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.matchevaluator?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("EEFD4DA6-24CD-48BB-8C32-BB44733A06FA")]
    [Description("Represents the method that is called each time a regular expression match is found during a Replace method operation.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMatchEvaluatorSingleton
    {
        [Description("Assign the replace method to the MatchEvaluator delegate.")]
        MatchEvaluator Create(IMatchEvaluator matchEvaluator);

    }
}
