// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.matchcollection?view=netframework-4.8.1

using System;
using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("2E24D7A2-C42D-4EE3-A3F3-E14A25059093")]
    [Description("Represents the set of successful matches found by iteratively applying a regular expression pattern to the input string. The collection is immutable (read-only) and has no public constructor. The Matches(String) method returns a MatchCollection object.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMatchCollection
    {
        //Properties
        int Count
        {
            [Description("Gets the number of matches.")]
            get;
        }

        bool IsReadOnly
        {
            [Description("Gets a value that indicates whether the collection is read only..")]
            get;
        }

        bool IsSynchronized
        {
            [Description("Gets a value indicating whether access to the collection is synchronized (thread-safe).")]
            get;
        }

        Match this[int index]
        {
            [Description("Gets an individual member of the collection.")]
            get;
        }

        object SyncRoot
        {
            [Description("Gets an object that can be used to synchronize access to the collection.")]
            get;
        }

        [Description("Copies all the elements of the collection to the given array starting at the given index.")]
        void CopyTo([In][Out] ref object[] array, int index);

        //CopyTo2(Empty[], Int32)

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("Provides an enumerator that iterates through the collection.")]
        IEnumerator GetEnumerator();

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        int GetHashCode();

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Returns a string that represents the current object.\r\n\r\n(Inherited from Object)")]
        string ToString();

    }
}
