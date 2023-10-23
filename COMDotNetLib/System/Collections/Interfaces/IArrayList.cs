// https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8.1

using GSystem = global::System;
using GCollections = global::System.Collections;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("FA145F11-5B72-4169-A318-BE970A28F985")]
    [Description("Implements the IList interface using an array whose size is dynamically increased as required.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IArrayList
    {
        // Properties

        int Capacity
        {
            [Description("Gets or sets the number of elements that the ArrayList can contain.")]
            get;
            [Description("Gets or sets the number of elements that the ArrayList can contain.")]
            set;
        }

        int Count
        {
            [Description("Gets the number of elements actually contained in the ArrayList.")]
            get;
        }

        bool IsFixedSize
        {
            [Description("Gets a value indicating whether the ArrayList has a fixed size.")]
            get;
        }

        bool IsReadOnly
        {
            [Description("Gets a value indicating whether the ArrayList is read-only.")]
            get;
        }

        bool IsSynchronized
        {
            [Description("Gets a value indicating whether access to the ArrayList is synchronized (thread safe).")]
            get;
        }

        object this[int index]
        {
            [Description("Gets or sets the element at the specified index.")]
            get;
            [Description("Gets or sets the element at the specified index.")]
            set;
        }

        object SyncRoot
        {
            [Description("Gets an object that can be used to synchronize access to the ArrayList.")]
            get;
        }

        // Methods

        [Description("Adds an object to the end of the ArrayList.")]
        int Add(object value);

        [Description("Adds the elements of an ICollection to the end of the ArrayList.")]
        void AddRange(GCollections.IEnumerable collection);

        [Description("Searches the entire sorted ArrayList for an element using the default comparer and returns the zero-based index of the element.")]
        int BinarySearch(object item);

        [Description("Searches the entire sorted ArrayList for an element using the specified comparer and returns the zero-based index of the element.")]
        int BinarySearch2(object item, GCollections.IComparer comparer);

        [Description("Searches a range of elements in the sorted ArrayList for an element using the specified comparer and returns the zero-based index of the element.")]
        int BinarySearch3(int index, int count, object item, GCollections.IComparer comparer);

        [Description("Removes all elements from the ArrayList.")]
        void Clear();

        [Description("Creates a shallow copy of the ArrayList.")]
        object Clone();

        [Description("Determines whether an element is in the ArrayList.")]
        bool Contains(object value);

        [Description("Copies the entire ArrayList to a compatible one-dimensional Array, starting at the beginning of the target array.")]
        void CopyTo(Array array);

        [Description("Copies the entire ArrayList to a compatible one-dimensional Array, starting at the beginning of the target array.")]
        void CopyTo2(Array array, int arrayIndex);

        [Description("Copies a range of elements from the ArrayList to a compatible one-dimensional Array, starting at the specified index of the target array.")]
        void CopyTo3(int index, Array array, int arrayIndex, int count);

        [Description("Returns an enumerator that iterates through a collection.")]
        GCollections.IEnumerator GetEnumerator();

        [Description("Returns an enumerator for a range of elements in the ArrayList.")]
        GCollections.IEnumerator GetEnumerator2(int index, int count);

        [Description("Returns an ArrayList which represents a subset of the elements in the source ArrayList.")]
        ArrayList GetRange(int index, int count);

        [Description("Searches for the specified Object and returns the zero-based index of the first occurrence within the entire ArrayList.")]
        int IndexOf(object value);

        [Description("Searches for the specified Object and returns the zero-based index of the first occurrence within the range of elements in the ArrayList that extends from the specified index to the last element.")]
        int IndexOf2(object value, int startIndex);

        [Description("Searches for the specified Object and returns the zero-based index of the first occurrence within the range of elements in the ArrayList that starts at the specified index and contains the specified number of elements.")]
        int IndexOf3(object value, int startIndex, int count);

        [Description("Inserts an element into the ArrayList at the specified index.")]
        void Insert(int index, object value);

        [Description("Inserts the elements of a collection into the ArrayList at the specified index.")]
        void InsertRange(int index, GCollections.ICollection c);

        [Description("Searches for the specified Object and returns the zero-based index of the last occurrence within the entire ArrayList.")]
        int LastIndexOf(object value);

        [Description("Searches for the specified Object and returns the zero-based index of the last occurrence within the range of elements in the ArrayList that extends from the first element to the specified index.")]
        int LastIndexOf2(object value, int startIndex);

        [Description("Searches for the specified Object and returns the zero-based index of the last occurrence within the range of elements in the ArrayList that contains the specified number of elements and ends at the specified index.")]
        int LastIndexOf3(object value, int startIndex, int count);

        [Description("Removes the first occurrence of a specific object from the ArrayList.")]
        void Remove(object value);

        [Description("Removes the element at the specified index of the ArrayList.")]
        void RemoveAt(int index);

        [Description("Removes a range of elements from the ArrayList.")]
        void RemoveRange(int index, int count);

        [Description("Reverses the order of the elements in the entire ArrayList.")]
        void Reverse();

        [Description("Reverses the order of the elements in the specified range.")]
        void Reverse2(int index, int count);

        [Description("Copies the elements of a collection over a range of elements in the ArrayList.")]
        void SetRange(int index, GCollections.ICollection c);

        [Description("Sorts the elements in the entire ArrayList.")]
        void Sort();

        [Description("Sorts the elements in the entire ArrayList using the specified comparer.")]
        void Sort2(GCollections.IComparer comparer);

        [Description("Sorts the elements in a range of elements in ArrayList using the specified comparer.")]
        void Sort3(int index, int count, GCollections.IComparer comparer);

        [Description("Copies the elements of the ArrayList to a new Object array.")]
        Array ToArray();

        [Description("Copies the elements of the ArrayList to a new array of the specified element type.")]
        Array ToArray2(Type type);

        [Description("Sets the capacity to the actual number of elements in the ArrayList.")]
        void TrimToSize();

        // Added to fix issue assigning value types.
        [Description("Sets the element at the specified index.")]
        void SetItem(int index, object value);
    }
}
