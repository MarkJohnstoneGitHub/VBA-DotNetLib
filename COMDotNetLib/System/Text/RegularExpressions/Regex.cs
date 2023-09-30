// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;
using GRegularExpressions = global::System.Text.RegularExpressions;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("E491897C-B93E-4A28-9FD2-9EF3BD1AAA06")]
    [ProgId("DotNetLib.System.Text.RegularExpressions.Regex")]
    [Description("Represents an immutable regular expression.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRegex))]
    public class Regex : IRegex
    {
        private GRegularExpressions.Regex _regex;
        private TimeSpan _matchTimeout;

        private static TimeSpan _infiniteMatchTimeout = new TimeSpan(GRegularExpressions.Regex.InfiniteMatchTimeout);

        // Constructors
        internal Regex(GRegularExpressions.Regex regex)
        {
            _regex = regex;
        }

        public Regex(string pattern)
        {
            _regex = new GRegularExpressions.Regex(pattern);
        }

        public Regex(string pattern, RegexOptions options)
        {
            _regex = new GRegularExpressions.Regex(pattern, (GRegularExpressions.RegexOptions)options);
        }

        public Regex(string pattern, RegexOptions options, TimeSpan matchTimeout)
        {
            _regex = new GRegularExpressions.Regex(pattern, (GRegularExpressions.RegexOptions)options, matchTimeout.WrappedTimeSpan);
        }

        // Fields

        public static TimeSpan InfiniteMatchTimeout => _infiniteMatchTimeout;

        // Properties

        internal GRegularExpressions.Regex WrappedRegex
        {
            get { return _regex; }
            set
            {
                _regex = value;
                _matchTimeout = new TimeSpan(value.MatchTimeout);
            }
        }

        public static int CacheSize
        {
            get => GRegularExpressions.Regex.CacheSize;
            set => GRegularExpressions.Regex.CacheSize = value;
        }

        public TimeSpan MatchTimeout => _matchTimeout;

        public RegexOptions Options => (RegexOptions)_regex.Options;

        public bool RightToLeft => _regex.RightToLeft;


        // Methods
        public static string Escape(string str)
        {
            return GRegularExpressions.Regex.Escape(str);
        }

        public string[] GetGroupNames()
        {
            return _regex.GetGroupNames();
        }

        public int[] GetGroupNumbers()
        {
            return _regex.GetGroupNumbers();
        }

        public string GroupNameFromNumber(int i)
        {
            return _regex.GroupNameFromNumber(i);
        }

        public int GroupNumberFromName(string name)
        {
            return _regex.GroupNumberFromName(name);
        }

        public bool IsMatch(string input)
        {
            return _regex.IsMatch(input);
        }

        public bool IsMatch2(string input, int startat)
        {
            return _regex.IsMatch(input, startat);
        }

        public static bool IsMatch(string input, string pattern, RegexOptions options)
        {
            return GRegularExpressions.Regex.IsMatch(input, pattern, (GRegularExpressions.RegexOptions)options);
        }

        public static bool IsMatch(string input, string pattern, RegexOptions options, TimeSpan matchTimeout)
        {
            return GRegularExpressions.Regex.IsMatch(input, pattern, (GRegularExpressions.RegexOptions)options, matchTimeout.WrappedTimeSpan);
        }

        public Match Match(string input)
        {
            return new Match(_regex.Match(input), this);
        }

        public Match Match2(string input, int startat)
        {
            return new Match(_regex.Match(input, startat), this);
        }

        public Match Match3(string input, int beginning, int length)
        {
            return new Match(_regex.Match(input, beginning, length), this);
        }

        public static Match Match(string input, string pattern)
        {
            return new Match(GRegularExpressions.Regex.Match(input, pattern));
        }

        public static Match Match(string input, string pattern, RegexOptions options)
        {
            return new Match(GRegularExpressions.Regex.Match(input, pattern, (GRegularExpressions.RegexOptions)options));
        }

        public static Match Match(string input, string pattern, RegexOptions options, TimeSpan matchTimeout)
        {
            return new Match(GRegularExpressions.Regex.Match(input, pattern, (GRegularExpressions.RegexOptions)options, matchTimeout.WrappedTimeSpan));
        }

        public MatchCollection Matches(string input)
        {
            return new MatchCollection(_regex.Matches(input),this);
        }

        public MatchCollection Matches(string input, int startat)
        {
            return new MatchCollection(_regex.Matches(input, startat), this);
        }

        public static MatchCollection Matches(string input, string pattern)
        {
            return new MatchCollection(GRegularExpressions.Regex.Matches(input, pattern));
        }

        public static MatchCollection Matches(string input, string pattern, RegexOptions options)
        {
            return new MatchCollection(GRegularExpressions.Regex.Matches(input, pattern, (GRegularExpressions.RegexOptions)options));
        }

        public static MatchCollection Matches(string input, string pattern, RegexOptions options, TimeSpan matchTimeout)
        {
            return new MatchCollection(GRegularExpressions.Regex.Matches(input, pattern, (GRegularExpressions.RegexOptions)options, matchTimeout.WrappedTimeSpan));
        }

        public string Replace(string input, string replacement)
        {
            return _regex.Replace(input, replacement);
        }

        public string Replace2(string input, string replacement, int count, int startat)
        {
            return _regex.Replace(input, replacement, count, startat);
        }

        public string Replace(string input, GRegularExpressions.MatchEvaluator evaluator)
        {
            return _regex.Replace(input, evaluator);
        }

        public string Replace(string input, GRegularExpressions.MatchEvaluator evaluator, int count)
        {
            return _regex.Replace(input, evaluator, count);
        }

        public string Replace(string input, GRegularExpressions.MatchEvaluator evaluator, int count, int startat)
        {
            return _regex.Replace(input, evaluator, count, startat);
        }

        public static string Replace(string input, string pattern, string replacement)
        {
            return GRegularExpressions.Regex.Replace(input, pattern, replacement);
        }

        public static string Replace(string input, string pattern, string replacement, RegexOptions options)
        {
            return GRegularExpressions.Regex.Replace(input, pattern, replacement, (GRegularExpressions.RegexOptions)options);
        }

        public static string Replace(string input, string pattern, string replacement, RegexOptions options, TimeSpan matchTimeout)
        {
            return GRegularExpressions.Regex.Replace(input, pattern, replacement, (GRegularExpressions.RegexOptions)options, matchTimeout.WrappedTimeSpan);
        }

        public static string Replace(string input, string pattern, GRegularExpressions.MatchEvaluator evaluator)
        {
            return GRegularExpressions.Regex.Replace(input, pattern, evaluator);
        }

        public static string Replace(string input, string pattern, GRegularExpressions.MatchEvaluator evaluator, RegexOptions options)
        {
            return GRegularExpressions.Regex.Replace(input, pattern, evaluator, (GRegularExpressions.RegexOptions)options);
        }

        public static string Replace(string input, string pattern, GRegularExpressions.MatchEvaluator evaluator, RegexOptions options, TimeSpan matchTimeout)
        {
            return GRegularExpressions.Regex.Replace(input, pattern, evaluator, (GRegularExpressions.RegexOptions)options, matchTimeout.WrappedTimeSpan);
        }

        //public string[] Split(string input)
        //{
        //    return _regex.Split(input);
        //}

        public string[] Split(string input, int count = 0)
        {
            return _regex.Split(input, count);
        }

        public string[] Split(string input, int count, int startat)
        {
            return _regex.Split(input, count, startat);
        }

        public static string[] Split(string input, string pattern)
        {
            return GRegularExpressions.Regex.Split(input, pattern);
        }
        public static string[] Split(string input, string pattern, RegexOptions options)
        {
            return GRegularExpressions.Regex.Split(input, pattern, (GRegularExpressions.RegexOptions)options);
        }

        public static string[] Split(string input, string pattern, RegexOptions options, TimeSpan matchTimeout)
        {
            return GRegularExpressions.Regex.Split(input, pattern, (GRegularExpressions.RegexOptions)options, matchTimeout.WrappedTimeSpan);
        }

        public override string ToString()
        { 
            return _regex.ToString(); 
        }

        public static string Unescape(string str)
        {
            return GRegularExpressions.Regex.Unescape(str);
        }

    }
}
