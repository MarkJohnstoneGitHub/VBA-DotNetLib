// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.capturecollection?view=netframework-4.8.1

using System;
using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("2D783821-4567-420F-ADD0-ED7024CAAD57")]
    [Description("Represents the set of captures made by a single capturing group. The collection is immutable (read-only) and has no public constructor.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICaptureCollection
    {
        //Properties
        int Count
        {
            [Description("Gets the number of substrings captured by the group.")]
            get;
        }

        bool IsReadOnly
        {
            [Description("Gets a value that indicates whether the collection is read only.")]
            get;
        }

        bool IsSynchronized
        {
            [Description("Gets a value that indicates whether access to the collection is synchronized (thread-safe).")]
            get;
        }

        Capture this[int i] 
        {
            [Description("Gets an individual member of the collection.")]
            get;
        }


        object SyncRoot
        {
            [Description("Gets an object that can be used to synchronize access to the collection.")]
            get;
        }

        [Description("Copies all the elements of the collection to the given array beginning at the given index.")]
        void CopyTo([In][Out] ref object[] array, int index);

        //CopyTo(Match[], Int32)

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
