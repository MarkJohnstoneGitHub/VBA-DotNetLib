using GCollections = global::System.Collections;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("9FCFB6D2-77DB-4416-BB5C-0BA5D54B4711")]
    [Description("Provides methods for creating, manipulating, searching, and sorting arrays, thereby serving as the base class for all arrays in the common language runtime.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IArray
    {
        // Properties

        bool IsFixedSize
        {
            [Description("Gets a value indicating whether the Array has a fixed size.")]
            get;
        }

        bool IsReadOnly
        {
            [Description("Gets a value indicating whether the Array is read-only.")]
            get;
        }

        bool IsSynchronized
        {
            [Description("Gets a value indicating whether access to the Array is synchronized (thread safe).")]
            get;
        }

        int Length
        {
            [Description("Gets the total number of elements in all the dimensions of the Array.")]
            get;
        }

        long LongLength
        {
            [Description("Gets a 64-bit integer that represents the total number of elements in all the dimensions of the Array.")]
            get;
        }

        int Rank
        {
            [Description("Gets the rank (number of dimensions) of the Array. For example, a one-dimensional array returns 1, a two-dimensional array returns 2, and so on.")]
            get;
        }

        object SyncRoot
        {
            [Description("Gets an object that can be used to synchronize access to the Array.")]
            get;
        }

        // Methods

        [Description("Creates a shallow copy of the Array.")]
        object Clone();

        [Description("Copies all the elements of the current one-dimensional array to the specified one-dimensional array starting at the specified destination array index. The index is specified as a 32-bit integer.")]
        void CopyTo(Array array, int index);

        [Description("Copies all the elements of the current one-dimensional array to the specified one-dimensional array starting at the specified destination array index. The index is specified as a 64-bit integer.")]
        void CopyTo(Array array, long index);

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("Returns an IEnumerator for the Array.")]
        GCollections.IEnumerator GetEnumerator();

        [Description("Serves as the default hash function.")]
        int GetHashCode();

        [Description("Gets a 32-bit integer that represents the number of elements in the specified dimension of the Array.")]
        int GetLength(int dimension);

        [Description("Gets a 64-bit integer that represents the number of elements in the specified dimension of the Array.")]
        long GetLongLength(int dimension);

        [Description("Gets the index of the first element of the specified dimension in the array.")]
        int GetLowerBound(int dimension);

        [Description("Gets the Type of the current instance.")]
        Type GetType();

        [Description("Gets the index of the last element of the specified dimension in the array.")]
        int GetUpperBound(int dimension);

        [Description("Gets the value at the specified position in the one-dimensional Array. The index is specified as a 32-bit integer.")]
        object GetValue(int index);

        [Description("Gets the value at the specified position in the two-dimensional Array. The indexes are specified as 32-bit integers.")]
        object GetValue(int index1, int index2);

        [Description("Gets the value at the specified position in the three-dimensional Array. The indexes are specified as 32-bit integers.")]
        object GetValue(int index1, int index2, int index3);

        [Description("Gets the value at the specified position in the multidimensional Array. The indexes are specified as an array of 32-bit integers.")]
        object GetValue([In] ref int[] indices);

        [Description("Gets the value at the specified position in the one-dimensional Array. The index is specified as a 64-bit integer.")] 
        object GetValue(long index);

        [Description("Gets the value at the specified position in the two-dimensional Array. The indexes are specified as 64-bit integers.")]
        object GetValue(long index1, long index2);

        [Description("Gets the value at the specified position in the three-dimensional Array. The indexes are specified as 64-bit integers.")]
        object GetValue(long index1, long index2, long index3);

        [Description("Gets the value at the specified position in the multidimensional Array. The indexes are specified as an array of 64-bit integers.")]
        object GetValue([In] ref long[] indices);

        [Description("Initializes every element of the value-type Array by calling the parameterless constructor of the value type.")]
        void Initialize();

        [Description("Sets a value to the element at the specified position in the one-dimensional Array. The index is specified as a 32-bit integer.")]
        void SetValue(object value, int index);

        [Description("Sets a value to the element at the specified position in the two-dimensional Array. The indexes are specified as 32-bit integers.")]
        void SetValue(object value, int index1, int index2);

        [Description("Sets a value to the element at the specified position in the three-dimensional Array. The indexes are specified as 32-bit integers.")]
        void SetValue(object value, int index1, int index2, int index3);

        [Description("Sets a value to the element at the specified position in the multidimensional Array. The indexes are specified as an array of 32-bit integers.")]
        void SetValue(object value, [In] ref int[] indices);

        [Description("Sets a value to the element at the specified position in the one-dimensional Array. The index is specified as a 64-bit integer.")]
        void SetValue(object value, long index);

        [Description("Sets a value to the element at the specified position in the two-dimensional Array. The indexes are specified as 64-bit integers.")]
        void SetValue(object value, long index1, long index2);

        [Description("Sets a value to the element at the specified position in the three-dimensional Array. The indexes are specified as 64-bit integers.")]
        void SetValue(object value, long index1, long index2, long index3);

        [Description("Sets a value to the element at the specified position in the multidimensional Array. The indexes are specified as an array of 64-bit integers.")]
        void SetValue(object value, [In] ref long[] indices);


        // Explicit Interface Implementations
        // https://learn.microsoft.com/en-us/dotnet/api/system.array?view=netframework-4.8.1#explicit-interface-implementations

        int Count
        {
            [Description("Gets the number of elements contained in the Array.")]
            get;
        }

        object this[int index]
        {
            [Description("Gets or sets the element at the specified index.")]
            get;
            [Description("Gets or sets the element at the specified index.")]
            set;
        }

        [Description("Calling this method always throws a NotSupportedException exception.")]
        int Add(object value);

        [Description("Removes all items from the IList.")]
        void Clear();

        [Description("Determines whether an element is in the IList.")]
        bool Contains(object value);

        [Description("Determines the index of a specific item in the IList.")]
        int IndexOf(object value);

        [Description("Inserts an item to the IList at the specified index.")]
        void Insert(int index, object value);

        [Description("Removes the first occurrence of a specific object from the IList.")]
        void Remove(object value);

        [Description("Removes the IList item at the specified index.")]
        void RemoveAt(int index);

        [Description("Determines whether the current collection object precedes, occurs in the same position as, or follows another object in the sort order.")]
        int CompareTo(object other, GCollections.IComparer comparer);

        [Description("Determines whether an object is equal to the current instance.")]
        bool Equals (object other, GCollections.IEqualityComparer comparer);

        [Description("Returns a hash code for the current instance.")]
        int GetHashCode(GCollections.IEqualityComparer comparer);

    }
}
