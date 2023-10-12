// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.match?view=netframework-4.8.1

using GRegularExpressions = global::System.Text.RegularExpressions;
using DotNetLib.Extensions;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("B77F425E-7914-42A4-B0EF-678DD0413241")]
    [ProgId("DotNetLib.System.Text.RegularExpressions.Match")]
    [Description("Represents the results from a single regular expression match.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMatch))]
    public class Match : Group, IMatch, IGroup, ICapture, IWrappedObject
    {
        internal Regex _regex;
        internal GRegularExpressions.Match _match;
        internal GroupCollection _groupcoll;

        // Constructors
        public Match(GRegularExpressions.Match match, Regex regex) : base(match)
        {
            _match = match;
            _regex = regex;
        }

        public Match(GRegularExpressions.Match match) : base(match)
        {
            _match = match;
        }

        //Properties
        internal GRegularExpressions.Match WrappedMatch
        {
            get { return _match; }
        }

        new public object WrappedObject => _match;

        public static Match Empty => new Match(GRegularExpressions.Match.Empty);

        public virtual GroupCollection Groups
        {
            get
            {
                if (_groupcoll == null)
                {
                    _groupcoll = new GroupCollection(this);
                }

                return _groupcoll;
            }
        }

        // Methods
        public Match NextMatch()
        {
            return  new Match(_match.NextMatch(),_regex);
        }

        public string Result(string replacement)
        { 
            return _match.Result(replacement); 
        }

        public static Match Synchronized(Match inner)
        {
            return new Match(GRegularExpressions.Match.Synchronized(inner.WrappedMatch));
        }

        //public bool Equals(object obj)
        //{

        //}

        public new virtual int GetHashCode()
        { 
            return _match.GetHashCode();
        }

        public override string ToString()
        { 
            return _match.ToString(); 
        }

        public new Type GetType()
        {
            return new Type(typeof(Match));
        }
    }
}
