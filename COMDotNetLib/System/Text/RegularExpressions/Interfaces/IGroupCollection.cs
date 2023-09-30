// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.groupcollection?view=netframework-4.8.1
using System;
using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("017AF829-E15F-4BE1-931E-353F5751B50D")]
    [Description("Returns the set of captured groups in a single match. The collection is immutable (read-only) and has no public constructor.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IGroupCollection
    {
        //Properties
        int Count
        {
            [Description("Returns the number of groups in the collection.")]
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

        Group this[int index]
        {
            [Description("Gets an individual member of the collection.")]
            get;
        }

        Group this[string groupname]
        {
            [Description("Enables access to a member of the collection by string index.")]
            get;
        }

        object SyncRoot
        {
            [Description("Gets an object that can be used to synchronize access to the collection.")]
            get;
        }

        [Description("Copies all the elements of the collection to the given array starting at the given index.")]
        void CopyTo([In][Out] ref object[] array, int index);

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
