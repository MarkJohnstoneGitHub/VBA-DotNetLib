// https://learn.microsoft.com/en-us/dotnet/api/system.collections.stack?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("7B9E50B3-A49C-4DAF-AC20-E03FF16384C7")]
    [Description("Represents a simple last-in-first-out (LIFO) non-generic collection of objects.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStack
    {
        int Count
        {
            [Description("Gets the number of elements contained in the Stack.")]
            get;
        }

        bool IsSynchronized
        {
            [Description("Gets a value indicating whether access to the Stack is synchronized (thread safe).")]
            get;
        }

        object SyncRoot
        {
            [Description("Gets an object that can be used to synchronize access to the Stack.")]
            get;
        }

        [Description("Removes all objects from the Stack.")]
        void Clear();

        [Description("Creates a shallow copy of the Stack.")]
        object Clone();

        [Description("Determines whether an element is in the Stack.")]
        bool Contains(object obj);

        [Description("Copies the Stack to an existing one-dimensional Array, starting at the specified array index.")]
        void CopyTo(Array array, int index);

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("Returns an IEnumerator for the Stack.")]
        GCollections.IEnumerator GetEnumerator();

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        int GetHashCode();

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Returns the object at the top of the Stack without removing it.")]
        object Peek();

        [Description("Removes and returns the object at the top of the Stack.")]
        object Pop();

        [Description("Inserts an object at the top of the Stack.")]
        void Push(object obj);

        [Description("Copies the Stack to a new array.")]
        Array ToArray();

        [Description("Copies the Stack to a new safearray.")]
        object[] ToSafeArray();

        [Description("Returns a string that represents the current object.\r\n\r\n(Inherited from Object)")]
        string ToString();


    }
}
