// https://learn.microsoft.com/en-us/dotnet/api/system.collections.queue?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("7203BD68-6902-4D0D-A219-0A9097D2045A")]
    [Description("Represents a first-in, first-out collection of objects.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IQueue
    {
        int Count
        {
            [Description("Gets the number of elements contained in the Queue.")]
            get; 
        }

        bool IsSynchronized 
        {
            [Description("Gets a value indicating whether access to the Queue is synchronized (thread safe).")]
            get;
        }

        object SyncRoot 
        {
            [Description("Gets an object that can be used to synchronize access to the Queue.")]
            get;
        }

        [Description("Removes all objects from the Queue.")]
        void Clear();

        [Description("Creates a shallow copy of the Queue.")]
        object Clone();

        [Description("Determines whether an element is in the Queue.")]
        bool Contains(object obj);

        [Description("Copies the Queue elements to an existing one-dimensional Array, starting at the specified array index.")]
        void CopyTo(Array array, int index);

        [Description("Removes and returns the object at the beginning of the Queue.")]
        object Dequeue();

        [Description("Adds an object to the end of the Queue.")]
        void Enqueue(object obj);

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("Returns an enumerator that iterates through the Queue.")]
        GCollections.IEnumerator GetEnumerator();

        [Description("Returns the object at the beginning of the Queue without removing it.")]
        object Peek();

        [Description("Copies the Queue elements to a new array.")]
        object[] ToArray();

        [Description("Sets the capacity to the actual number of elements in the Queue.")]
        void TrimToSize();
    }
}
