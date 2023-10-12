// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.matchcollection?view=netframework-4.8.1

using GRegularExpressions = global::System.Text.RegularExpressions;
using GSystem = global::System;
using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("6402A435-4D50-4AE5-B743-9DC79DF3A803")]
    [ProgId("DotNetLib.System.Text.RegularExpressions.MatchCollection")]
    [Description("Represents the set of successful matches found by iteratively applying a regular expression pattern to the input string. The collection is immutable (read-only) and has no public constructor. The Matches(String) method returns a MatchCollection object.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMatchCollection))]
    public class MatchCollection : ICollection, IEnumerable, IMatchCollection
    {
        internal Regex _regex;
        private GRegularExpressions.MatchCollection _matchCollection;
        private ArrayList _matches;

        // Constructors
        internal MatchCollection(GRegularExpressions.MatchCollection matchCollection, Regex regex)
        {
            _matchCollection = matchCollection;
            _regex = regex;
            _matches = WrapMatchCollection();
        }

        internal MatchCollection(GRegularExpressions.MatchCollection matchCollection)
        {
            _matchCollection = matchCollection;
            _matches = WrapMatchCollection();
        }

        public int Count => _matchCollection.Count;

        public bool IsReadOnly => _matchCollection.IsReadOnly;

        public object SyncRoot => _matches.SyncRoot;

        public bool IsSynchronized => _matches.IsSynchronized;


        public void CopyTo([In, Out] ref object[] array, int index)
        {
            _matches.CopyTo(array, index);
        }

        public void CopyTo(Array array, int index)
        {
            _matches.CopyTo(array.WrappedArray , index);
        }

        public void CopyTo(GSystem.Array array, int index)
        {
            ((ICollection)_matchCollection).CopyTo(array, index);
        }


        public IEnumerator GetEnumerator()
        {
            return _matches.GetEnumerator();
        }


        //
        // Summary:
        //     Gets an individual member of the collection.
        //
        // Parameters:
        //   i:
        //     Index into the System.Text.RegularExpressions.Match collection.
        //
        // Returns:
        //     The captured substring at position i in the collection.
        //
        // Exceptions:
        //   T:System.ArgumentOutOfRangeException:
        //     i is less than 0 or greater than or equal to System.Text.RegularExpressions.MatchCollection.Count.
        //
        //   T:System.Text.RegularExpressions.RegexMatchTimeoutException:
        //     A time-out occurred.
        public virtual Match this[int i]
        {
            get
            {
                return (Match)_matches[i];
            }
        }

        internal GSystem.Collections.ArrayList WrapMatchCollection()
        {
            GSystem.Collections.ArrayList matchCollection = new GSystem.Collections.ArrayList(_matchCollection.Count);
            foreach (GRegularExpressions.Match match in _matchCollection)
            {
                matchCollection.Add(new Match(match));
            }
            return matchCollection;
        }

        public new Type GetType()
        {
            return new Type(typeof(MatchCollection));
        }
    }
}
