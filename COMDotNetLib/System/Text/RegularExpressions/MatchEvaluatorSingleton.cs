// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.matchevaluator?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("05CFC860-82B5-429E-9CF2-EC67DF674E21")]
    [ProgId("DotNetLib.System.Text.RegularExpressions.MatchEvaluatorSingleton")]
    [Description("Represents the method that is called each time a regular expression match is found during a Replace method operation.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMatchEvaluatorSingleton))]
    public class MatchEvaluatorSingleton : IMatchEvaluatorSingleton
    {
        public MatchEvaluatorSingleton() { }

        // Factory Methods
        public MatchEvaluator Create(IMatchEvaluator matchEvaluator)
        {
            return new MatchEvaluator(matchEvaluator);
        }


    }
}
