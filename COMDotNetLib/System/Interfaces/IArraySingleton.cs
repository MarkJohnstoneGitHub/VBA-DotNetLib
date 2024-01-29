// https://learn.microsoft.com/en-us/dotnet/api/system.array?view=netframework-4.8.1

using GSystem = global::System;
using GCollections = global::System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("1CC4D6F2-35AC-44F7-A666-D1845A1E30D0")]
    [Description("Provides methods for creating, manipulating, searching, and sorting arrays, thereby serving as the base class for all arrays in the common language runtime.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IArraySingleton
    {
        // Note added to wrap a mscorlib.Array for mscorlib.ICollection void CopyTo2(Array array, int index);
        [Description("Creates an Array by wrapping an mscorlib.Array")]
        Array Create(GSystem.Array array);

        [Description("Searches an entire one-dimensional sorted array for a specific element, using the IComparable interface implemented by each element of the array and by the specified object.")]
        int BinarySearch(Array array, object value);

        [Description("Searches an entire one-dimensional sorted array for a value using the specified IComparer interface.")]
        int BinarySearch2(Array array, object value, GCollections.IComparer comparer = null);

        [Description("Searches a range of elements in a one-dimensional sorted array for a value, using the specified IComparer interface.")]
        int BinarySearch3(Array array, int index, int length, object value, GCollections.IComparer comparer = null);

        //[Description("Searches an entire one-dimensional sorted array for a specific element, using the IComparable interface implemented by each element of the array and by the specified object.")]
        //int BinarySearch2(Array array, object value);

        //[Description("Searches a range of elements in a one-dimensional sorted array for a value, using the IComparable interface implemented by each element of the array and by the specified value.")]
        //int BinarySearch2(Array array, int index, int length, object value);

        [Description("Sets a range of elements in an array to the default value of each element type.")]
        void Clear(Array array, int index, int length);

        [Description("Copies a range of elements from an Array starting at the specified source index and pastes them to another Array starting at the specified destination index. Guarantees that all changes are undone if the copy does not succeed completely.")]
        void ConstrainedCopy(Array sourceArray, int sourceIndex, Array destinationArray, int destinationIndex, int length);

        [Description("Copies a range of elements from an Array starting at the first element and pastes them into another Array starting at the first element. The length is specified as a 32-bit integer.")]
        void Copy(Array sourceArray, Array destinationArray, int length);

        [Description("Copies a range of elements from an Array starting at the first element and pastes them into another Array starting at the first element. The length is specified as a 64-bit integer.")]
        void Copy2(Array sourceArray, Array destinationArray, long length);

        [Description("Copies a range of elements from an Array starting at the specified source index and pastes them to another Array starting at the specified destination index. The length and the indexes are specified as 32-bit integers.")]
        void Copy3(Array sourceArray, int sourceIndex, Array destinationArray, int destinationIndex, int length);

        [Description("Creates a one-dimensional Array of the specified Type and length, with zero-based indexing.")]
        Array CreateInstance(Type elementType, int length);
        
        [Description("Creates a two-dimensional Array of the specified Type and dimension lengths, with zero-based indexing.")]
        Array CreateInstance2(Type elementType, int length1, int length2);

        [Description("Creates a three-dimensional Array of the specified Type and dimension lengths, with zero-based indexing.")]
        Array CreateInstance3(Type elementType, int length1, int length2, int length3);

        [Description("Creates a multidimensional Array of the specified Type and dimension lengths, with zero-based indexing. The dimension lengths are specified in an array of 32-bit integers.")]
        Array CreateInstance4(Type elementType, [In] ref int[] lengths);

        [Description("Creates a multidimensional Array of the specified Type and dimension lengths, with zero-based indexing. The dimension lengths are specified in an array of 64-bit integers.")]
        Array CreateInstance5(Type elementType, [In] ref long[] lengths);

        [Description("Creates a multidimensional Array of the specified Type and dimension lengths, with the specified lower bounds.")]
        Array CreateInstance6(Type elementType, [In] ref int[] lengths, [In] ref int[] lowerBounds);

        [Description("Searches for the specified object and returns the index of its first occurrence in a one-dimensional array.")]
        int IndexOf(Array array, object value);

        [Description("Searches for the specified object in a range of elements of a one-dimensional array, and returns the index of its first occurrence. The range extends from a specified index to the end of the array.")]
        int IndexOf2(Array array, object value, int startIndex);

        [Description("Searches for the specified object in a range of elements of a one-dimensional array, and returns the index of ifs first occurrence. The range extends from a specified index for a specified number of elements.")]
        int IndexOf3(Array array, object value, int startIndex, int count);

        [Description("Searches for the specified object and returns the index of the last occurrence within the entire one-dimensional Array.")]
        int LastIndexOf(Array array, object value);

        [Description("Searches for the specified object and returns the index of the last occurrence within the range of elements in the one-dimensional Array that extends from the first element to the specified index.")]
        int LastIndexOf2(Array array, object value, int startIndex);

        [Description("Searches for the specified object and returns the index of the last occurrence within the range of elements in the one-dimensional Array that contains the specified number of elements and ends at the specified index.")]
        int LastIndexOf3(Array array, object value, int startIndex, int count);

        [Description("Changes the number of elements of a one-dimensional array to the specified new size.")]
        void Resize(ref Array array, int newSize);

        [Description("Reverses the sequence of the elements in the entire one-dimensional Array.")]
        void Reverse(Array array);

        [Description("Reverses the sequence of a subset of the elements in the one-dimensional Array.")]
        void Reverse2(Array array, int index, int length);

        //[Description("Sorts the elements in an entire one-dimensional Array using the IComparable implementation of each element of the Array.")]
        //void Sort2(Array array);

        [Description("Sorts the elements in a one-dimensional Array using the specified IComparer.")]
        void Sort(Array array, GCollections.IComparer comparer = null);

        //[Description("")]
        //void Sort2(Array array, int index, int length);

        [Description("Sorts the elements in a range of elements in a one-dimensional Array using the specified IComparer.")]
        void Sort2(Array array, int index, int length, GCollections.IComparer comparer = null);

        //[Description("Sorts a pair of one-dimensional Array objects (one contains the keys and the other contains the corresponding items) based on the keys in the first Array using the IComparable implementation of each key.")]
        //void Sort2(Array keys, Array items);

        [Description("Sorts a pair of one-dimensional Array objects (one contains the keys and the other contains the corresponding items) based on the keys in the first Array using the specified IComparer.")]
        void Sort3(Array keys, Array items, GCollections.IComparer comparer = null);

        //[Description("Sorts a range of elements in a pair of one-dimensional Array objects (one contains the keys and the other contains the corresponding items) based on the keys in the first Array using the IComparable implementation of each key.")]
        //void Sort2(Array keys, Array items, int index, int length);

        [Description("Sorts a range of elements in a pair of one-dimensional Array objects (one contains the keys and the other contains the corresponding items) based on the keys in the first Array using the specified IComparer.")]
        void Sort4(Array keys, Array items, int index, int length, GCollections.IComparer comparer = null);

        // Extension Methods
        // https://learn.microsoft.com/en-us/dotnet/api/system.array?view=netframework-4.8.1#extension-methods

        //Todo : Issue must be implemented on a non-generic static class
        //[Description("Converts an IEnumerable to an IQueryable.")]
        //GSystem.Linq.IQueryable AsQueryable(this GCollections.IEnumerable source);

    }
}
