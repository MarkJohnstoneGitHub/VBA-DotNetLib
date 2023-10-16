// https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("4249C35B-BF20-4E1E-B916-995C25063F82")]
    [Description("Represents a collection of key/value pairs that are sorted by the keys and are accessible by key and by index.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISortedList
    {
        // Properties
        int Capacity
        {
            [Description("Gets or sets the capacity of a SortedList object.")]
            get;
            [Description("Gets or sets the capacity of a SortedList object.")]
            set;
        }

        int Count
        {
            [Description("Gets the number of elements contained in a SortedList object.")]
            get;
        }

        bool IsFixedSize
        {
            [Description("Gets a value indicating whether a SortedList object has a fixed size.")]
            get;
        }

        bool IsReadOnly
        {
            [Description("Gets a value indicating whether a SortedList object is read-only.")]
            get;
        }

        bool IsSynchronized
        {
            [Description("Gets a value indicating whether access to a SortedList object is synchronized (thread safe).")]
            get;
        }

        object this[object key]
        {
            [Description("Gets or sets the value associated with a specific key in a SortedList object.")]
            get;
            [Description("Gets or sets the value associated with a specific key in a SortedList object.")]
            set;
        }

        ICollection Keys 
        {
            [Description("Gets the keys in a SortedList object.")]
            get;
        }

        object SyncRoot 
        {
            [Description("Gets an object that can be used to synchronize access to a SortedList object.")]
            get;
        }

        ICollection Values 
        {
            [Description("Gets the values in a SortedList object.")]
            get;
        }

        // Methods

        [Description("Adds an element with the specified key and value to a SortedList object.")]
        void Add(object key, object value);

        [Description("Removes all elements from a SortedList object.")]
        void Clear();

        [Description("Creates a shallow copy of a SortedList object.")]
        object Clone();

        [Description("Determines whether a SortedList object contains a specific key.")]
        bool Contains(object key);

        [Description("Determines whether a SortedList object contains a specific key.")]
        bool ContainsKey(object key);

        [Description("Determines whether a SortedList object contains a specific value.")] 
        bool ContainsValue(object value);

        [Description("Copies SortedList elements to a one-dimensional Array object, starting at the specified index in the array.")]
        void CopyTo(Array array, int arrayIndex);

        [Description("Determines whether the specified object is equal to the current object.")]
        bool Equals(object obj);

        [Description("Gets the value at the specified index of a SortedList object.")]
        object GetByIndex(int index);

        [Description("Returns an IDictionaryEnumerator object that iterates through a SortedList object.")]
        GCollections.IDictionaryEnumerator GetEnumerator();

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        int GetHashCode();

        [Description("Gets the key at the specified index of a SortedList object.")]
        object GetKey(int index);

        [Description("Gets the keys in a SortedList object.")] 
        GCollections.IList GetKeyList();

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")] 
        Type GetType();

        [Description("Gets the values in a SortedList object.")] 
        GCollections.IList GetValueList();

        [Description("Returns the zero-based index of the specified key in a SortedList object.")]
        int IndexOfKey(object key);

        [Description("Returns the zero-based index of the first occurrence of the specified value in a SortedList object.")]
        int IndexOfValue(object value);

        [Description("Removes the element with the specified key from a SortedList object.")]
        void Remove(object key);

        [Description("Removes the element at the specified index of a SortedList object.")]
        void RemoveAt(int index);

        [Description("Replaces the value at a specific index in a SortedList object.")]
        void SetByIndex(int index, object value);

        [Description("Returns a string that represents the current object.")]
        string ToString();

        [Description("Sets the capacity to the actual number of elements in a SortedList object.")]
        void TrimToSize();


    }
}
