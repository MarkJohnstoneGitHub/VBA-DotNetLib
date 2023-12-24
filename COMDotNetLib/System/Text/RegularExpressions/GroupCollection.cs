// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.groupcollection?view=netframework-4.8.1

//Empty.Groups
using GSystem = global::System;
using GRegularExpressions = global::System.Text.RegularExpressions;
using System.Collections;
using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Linq;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("52ED5D22-3EEB-4DA0-8606-DB900ABC919B")]
    [ProgId("DotNetLib.System.Text.RegularExpressions.GroupCollection")]
    [Description("")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IGroupCollection))]
    public class GroupCollection : IGroupCollection, ICollection
    {
        internal Match _match;
        internal Group[] _groups;

        internal GroupCollection(Match match)
        {
            _match = match;
            _groups = WrapGroup();
        }
        
        public Group this[int index] =>  _groups[index];

        public Group this[string groupname] => new Group(_match.WrappedMatch.Groups[groupname]);


        public int Count => _groups.Count();

        //
        // Summary:
        //     Gets an object that can be used to synchronize access to the System.Text.RegularExpressions.GroupCollection.
        //
        // Returns:
        //     A copy of the System.Text.RegularExpressions.Empty object to synchronize.
        public object SyncRoot => _match;

        //
        // Summary:
        //     Gets a value that indicates whether access to the System.Text.RegularExpressions.GroupCollection
        //     is synchronized (thread-safe).
        //
        // Returns:
        //     false in all cases.
        public bool IsSynchronized => false;

        //
        // Summary:
        //     Gets a value that indicates whether the collection is read-only.
        //
        // Returns:
        //     true in all cases.
        public bool IsReadOnly => true;

        public void CopyTo(Array array, int index)
        {
            throw new NotImplementedException();
        }

        public void CopyTo([In, Out] ref object[] array, int index)
        {
            throw new NotImplementedException();
        }

        //Todo: Check implementation
        public IEnumerator GetEnumerator()
        {
            return  new GroupEnumerator(this);
        }

        internal Group[] WrapGroup()
        {
            Group[] groups = new Group[_match.WrappedMatch.Groups.Count];

            int  i = 0;
            foreach (GRegularExpressions.Group group in _match.WrappedMatch.Groups)
            {
                groups[i++] = new Group(group);
            }
            return groups;
        }

        public new Type GetType()
        {
            return new Type(typeof(GroupCollection));
        }

        public void CopyTo(global::System.Array array, int index)
        {
            _groups.CopyTo(array, index);
        }
    }

    /*
 * This non-public enumerator lists all the captures
 * Should it be public?
 */
    internal class GroupEnumerator : IEnumerator
    {
        internal GroupCollection _rgc;
        internal int _curindex;

        /*
         * Nonpublic constructor
         */
        internal GroupEnumerator(GroupCollection rgc)
        {
            _curindex = -1;
            _rgc = rgc;
        }

        /*
         * As required by IEnumerator
         */
        public bool MoveNext()
        {
            int size = _rgc.Count;

            if (_curindex >= size)
                return false;

            _curindex++;

            return (_curindex < size);
        }

        /*
         * As required by IEnumerator
         */
        public object Current
        {
            get { return Capture; }
        }

        /*
         * Returns the current capture
         */
        public Capture Capture
        {
            get
            {
                if (_curindex < 0 || _curindex >= _rgc.Count)
                    throw new InvalidOperationException("EnumNotStarted");

                return _rgc[_curindex];
            }
        }

        /*
         * Reset to before the first item
         */
        public void Reset()
        {
            _curindex = -1;
        }
    }
}
