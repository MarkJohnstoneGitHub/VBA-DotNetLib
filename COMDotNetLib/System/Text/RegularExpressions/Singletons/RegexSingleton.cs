// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?view=netframework-4.8.1

using GRegularExpressions = global::System.Text.RegularExpressions;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("E166B902-6153-4ADF-958F-04D599779A7C")]
    [ProgId("DotNetLib.System.Text.RegularExpressions.RegexSingleton")]
    [Description("Represents an immutable regular expression.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRegexSingleton))]
    public class RegexSingleton  : IRegexSingleton
    {
        public RegexSingleton() { }

        // Factory Methods

        public Regex Create(string pattern, RegexOptions options = RegexOptions.None)
        {
            return new Regex(pattern,options);
        }

        public Regex Create2(string pattern, RegexOptions options, TimeSpan matchTimeout)
        {
            return new Regex(pattern, options, matchTimeout);
        }

        // Properties
        public int CacheSize 
        { 
            get => Regex.CacheSize; 
            set => Regex.CacheSize = value;
        }

        public TimeSpan InfiniteMatchTimeout => Regex.InfiniteMatchTimeout;

        public string Escape(string str)
        {
            return Regex.Escape(str);
        }

        public bool IsMatch(string input, string pattern, RegexOptions options = RegexOptions.None)
        {
            return Regex.IsMatch(input, pattern, options);
        }

        public bool IsMatch2(string input, string pattern, RegexOptions options, TimeSpan matchTimeout)
        {
            return Regex.IsMatch(input, pattern, options, matchTimeout);
        }

        public Match Match(string input, string pattern, RegexOptions options = RegexOptions.None)
        {
            return Regex.Match(input,pattern, options);
        }

        public Match Match2(string input, string pattern, RegexOptions options, TimeSpan matchTimeout)
        {
            return Regex.Match(input, pattern, options,matchTimeout);
        }

        public MatchCollection Matches(string input, string pattern, RegexOptions options = RegexOptions.None)
        {
            return Regex.Matches(input, pattern, options);
        }

        public MatchCollection Matches2(string input, string pattern, RegexOptions options, TimeSpan matchTimeout)
        {
            return Regex.Matches(input, pattern, options, matchTimeout);
        }

        public string Replace(string input, string pattern, string replacement, RegexOptions options = RegexOptions.None)
        {
            return Regex.Replace(input, pattern, replacement, options);
        }

        public string Replace2(string input, string pattern, string replacement, RegexOptions options, TimeSpan matchTimeout)
        {
            return Regex.Replace(input, pattern, replacement, options, matchTimeout);
        }

        public string Replace3(string input, string pattern, MatchEvaluator evaluator, RegexOptions options = RegexOptions.None)
        {
            //GRegularExpressions.MatchEvaluator matchEvaluator = new GRegularExpressions.MatchEvaluator(evaluator.MatchEvaluatorCallBack);
            //return GRegularExpressions.Regex.Replace(input, pattern, matchEvaluator, (GRegularExpressions.RegexOptions)options);
            return GRegularExpressions.Regex.Replace(input, pattern, evaluator.Delegate, (GRegularExpressions.RegexOptions)options);
        }

        public string Replace4(string input, string pattern, MatchEvaluator evaluator, RegexOptions options, TimeSpan matchTimeout)
        {
            //GRegularExpressions.MatchEvaluator matchEvaluator = new GRegularExpressions.MatchEvaluator(evaluator.MatchEvaluatorCallBack);
            //return GRegularExpressions.Regex.Replace(input, pattern, matchEvaluator, (GRegularExpressions.RegexOptions)options, matchTimeout.WrappedTimeSpan);
            return GRegularExpressions.Regex.Replace(input, pattern, evaluator.Delegate, (GRegularExpressions.RegexOptions)options, matchTimeout.WrappedTimeSpan);
        }

        public string[] Split(string input, string pattern, RegexOptions options = RegexOptions.None)
        {
            return Regex.Split(input, pattern, options);
        }

        public string[] Split2(string input, string pattern, RegexOptions options, TimeSpan matchTimeout)
        {
            return Regex.Split(input, pattern, options, matchTimeout);
        }

        public string Unescape(string str)
        {
            return Regex.Unescape(str);
        }



    }
}
