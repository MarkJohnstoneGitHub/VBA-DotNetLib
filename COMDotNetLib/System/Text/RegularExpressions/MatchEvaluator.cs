// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.matchevaluator?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;
using GRegularExpressions = global::System.Text.RegularExpressions;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("877E5AB2-84CE-49AA-BB40-FAC72921C3BF")]
    [ProgId("DotNetLib.System.Text.RegularExpressions.MatchEvaluator")]
    [Description("Represents the method that is called each time a regular expression match is found during a Replace method operation.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMatchEvaluator))]
    public class MatchEvaluator : IMatchEvaluator
    {
        IMatchEvaluator _evaluator;

        public MatchEvaluator(IMatchEvaluator matchEvaluator)
        {
            _evaluator = matchEvaluator;
        }

        string IMatchEvaluator.MatchEvaluator(Match match)
        {
            return _evaluator.MatchEvaluator(match);
        }

        internal string RegexMatchEvaluator(GRegularExpressions.Match m)
        {
            Match myMatch = new Match(m);
            return _evaluator.MatchEvaluator(myMatch);
        }

    }
}
