// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

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

        public string Escape(string str)
        {
            return Regex.Escape(str);
        }

        public string Replace(string input, string pattern, string replacement, RegexOptions options = RegexOptions.None)
        {
            return Regex.Replace(input, pattern, replacement, options);
        }

        public string Replace2(string input, string pattern, string replacement, RegexOptions options, TimeSpan matchTimeout)
        {
            return Regex.Replace(input, pattern, replacement, options, matchTimeout.WrappedTimeSpan);
        }

        public string Unescape(string str)
        {
            return Regex.Unescape(str);
        }

    }
}
