using System.ComponentModel;
using System.Runtime.InteropServices;
using System;
using GSystem = global::System;

// https://learn.microsoft.com/en-us/dotnet/api/system.collections.ilist?view=netframework-4.8.1
namespace DotNetLib.System.Collections
{
    [Description("Represents a strongly typed list of objects that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [ComVisible(true)]
    [Guid("F03A9DAE-AF67-4C90-91C1-7EED79A37EF1")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]

    //
    // Summary:
    //     Represents a collection of objects that can be individually accessed by index.
    //
    // Type parameters:
    //   T:
    //     The type of elements in the list.

    public interface IList
    {
        // Constructors
        [Description("Initializes a new instance of the List < T > class that is empty and has the default initial capacity.")]
        List Create(object listType);

        [Description("Initializes a new instance of the List<T> class that is empty and has the specified initial capacity.")]
        List Create2(object listType, int capacity);

        [Description("Initializes a new instance of the List<T> class that contains elements copied from the specified collection and has sufficient capacity to accommodate the number of elements copied.")]
        List CreateFromIEnumerable(GSystem.Collections.IEnumerable collection);

        //
        // Summary:
        //     Gets or sets the element at the specified index.
        //
        // Parameters:
        //   index:
        //     The zero-based index of the element to get or set.
        //
        // Returns:
        //     The element at the specified index.
        //
        // Exceptions:
        //   T:System.ArgumentOutOfRangeException:
        //     index is not a valid index in the System.Collections.Generic.IList`1.
        //
        //   T:System.NotSupportedException:
        //     The property is set and the System.Collections.Generic.IList`1 is read-only.
        [Description("Gets or sets the element at the specified index.")]
        object this[int index]
        {
            get;
            set;
        }

        [Description("Adds an object to the end of the List<T>.")]
        void Add(object value);

        [Description("Adds the elements of the specified collection to the end of the List<T>.")]
        void AddRange(GSystem.Collections.IEnumerable collection);

        [Description("Removes all elements from the List<T>.")]
        void Clear();

        [Description("Determines whether an element is in the List<T>.")]
        bool Contains(object value);

        //System.Collections.Generic.List<TOutput> ConvertAll<TOutput>(Converter<T, TOutput> converter);

        [Description("Copies the entire List<T> to a compatible one-dimensional array, starting at the beginning of the target array.")]
        void CopyTo(object[] array);


        [Description("Copies the entire List<T> to a compatible one-dimensional array, starting at the specified index of the target array.")]
        void CopyTo2(object[] array, int arrayIndex);

        [Description("Copies a range of elements from the List<T> to a compatible one-dimensional array, starting at the specified index of the target array.")]
        void CopyTo3(int index, object[] array, int arrayIndex, int count);

        [Description("Searches for the specified object and returns the zero-based index of the first occurrence within the entire List<T>.")]
        int IndexOf(object item);

        [Description("Searches for the specified object and returns the zero-based index of the first occurrence within the range of elements in the List<T> that extends from the specified index to the last element.")]
        int IndexOf2(object item, int index);

        [Description("Searches for the specified object and returns the zero-based index of the first occurrence within the range of elements in the List<T> that starts at the specified index and contains the specified number of elements.")]
        int IndexOf3(object item, int index, int count);


        [Description("Inserts an element into the List<T> at the specified index.")]
        void Insert(int index, object item);

        [Description("Searches for the specified object and returns the zero-based index of the last occurrence within the entire List<T>.")]
        int LastIndexOf(object item);

        [Description("Searches for the specified object and returns the zero-based index of the last occurrence within the range of elements in the List<T> that extends from the first element to the specified index.")]
        int LastIndexOf2(object item, int index);

        [Description("Searches for the specified object and returns the zero-based index of the last occurrence within the range of elements in the List<T> that contains the specified number of elements and ends at the specified index.")]
        int LastIndexOf3(object item, int index, int count);

        [Description("Removes the first occurrence of a specific object from the IList.")] 
        bool Remove(object value);

        [Description("Removes the element at the specified index of the List<T>.")]
        void RemoveAt(int index);

        [Description("Removes a range of elements from the List<T>.")]
        void RemoveRange(int index, int count);

        [Description("Reverses the order of the elements in the entire List<T>.")]
        void Reverse();

        //public void Sort (Comparison<T> comparison);
        //public void Sort (int index, int count, System.Collections.Generic.IComparer<T> comparer);
        [Description("Sorts the elements in the entire List<T> using the default comparer.")]
        void Sort();

        [Description("Copies the elements of the List<T> to a new array.")]
        object[] ToArray();

        [Description("Sets the capacity to the actual number of elements in the List<T>, if that number is less than a threshold value.")]
        void TrimExcess();

        //[Description("Determines whether every element in the List<T> matches the conditions defined by the specified predicate.")]
        //bool TrueForAll(Predicate<T> match);
    }

}