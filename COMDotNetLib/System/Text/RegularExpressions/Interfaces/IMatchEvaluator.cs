// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.matchevaluator?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("B1F97124-C7FF-407E-A6BE-656EC2F9B6E0")]
    [Description("Represents the method that is called each time a regular expression match is found during a Replace method operation.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMatchEvaluator
    {
        [Description("Represents the method that is called each time a regular expression match is found during a Replace method operation.")]
        string MatchEvaluator(Match match);
    }
}

//public delegate string MatchEvaluator(Match match); // a type itself