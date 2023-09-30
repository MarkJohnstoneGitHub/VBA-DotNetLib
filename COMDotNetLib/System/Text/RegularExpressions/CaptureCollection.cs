// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.capturecollection?view=netframework-4.8.1

using GRegularExpressions = global::System.Text.RegularExpressions;
using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Linq;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("35D346F0-623C-41A4-AB67-DE5FADA8AF79")]
    [ProgId("DotNetLib.System.Text.RegularExpressions.CaptureCollection")]
    [Description("Represents the set of captures made by a single capturing group. The collection is immutable (read-only) and has no public constructor.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICaptureCollection))]
    public class CaptureCollection : ICaptureCollection, ICollection, IEnumerable
    {      
        internal Group _group;
        internal Capture[] _captures;

        // Constructors
        internal CaptureCollection(Group group)
        {
            _group = group;
            _captures = WrapCaptures();
        }

        // Properties

        public Capture this[int i] => _captures[i]; 


        //
        // Summary:
        //     Gets an object that can be used to synchronize access to the collection.
        //
        // Returns:
        //     An object that can be used to synchronize access to the collection.
        public object SyncRoot => _group;

        //
        // Summary:
        //     Gets a value that indicates whether access to the collection is synchronized
        //     (thread-safe).
        //
        // Returns:
        //     false in all cases.
        public bool IsSynchronized => false;

        //
        // Summary:
        //     Gets a value that indicates whether the collection is read only.
        //
        // Returns:
        //     true in all cases.
        public bool IsReadOnly => true;


        //
        // Summary:
        //     Gets the number of substrings captured by the group.
        //
        // Returns:
        //     The number of items in the System.Text.RegularExpressions.CaptureCollection.
        public int Count
        {
            get
            {
                return _captures.Count();
            }
        }

        public void CopyTo([In, Out] ref object[] array, int arrayIndex)
        {
            if (array == null)
            {
                throw new ArgumentNullException("array");
            }

            int num = arrayIndex;
            for (int i = 0; i < Count; i++)
            {
                array.SetValue(this[i], num);
                num++;
            }
        }

        public void CopyTo(Array array, int index)
        {
            throw new NotImplementedException();
        }


        // Todo check implementation
        public IEnumerator GetEnumerator()
        {
            return new CaptureEnumerator(this);
        }

        internal Capture[] WrapCaptures()
        {
            Capture[] captures = new Capture[_group.WrappedGroup.Captures.Count];

            int i=0;
            foreach (GRegularExpressions.Capture capture in _group.WrappedGroup.Captures)
            {
                captures[i++] = new Capture(capture);
            }
            return captures;
        }
    }


 /*
 * This non-public enumerator lists all the captures
 * Should it be public?
 */
#if !SILVERLIGHT
    [Serializable()]
#endif
    internal class CaptureEnumerator : IEnumerator
    {
        internal CaptureCollection _rcc;
        internal int _curindex;

        /*
         * Nonpublic constructor
         */
        internal CaptureEnumerator(CaptureCollection rcc)
        {
            _curindex = -1;
            _rcc = rcc;
        }

        /*
         * As required by IEnumerator
         */
        public bool MoveNext()
        {
            int size = _rcc.Count;

            if (_curindex >= size)
                return false;

            _curindex++;

            return (_curindex < size);
        }

        /*
         * As required by IEnumerator
         */
        public Object Current
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
                if (_curindex < 0 || _curindex >= _rcc.Count)
                    throw new InvalidOperationException("EnumNotStarted");

                return _rcc[_curindex];
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
