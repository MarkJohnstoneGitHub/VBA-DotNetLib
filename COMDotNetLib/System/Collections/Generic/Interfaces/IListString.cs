// https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System;
using System.Collections;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a strongly typed list of strings that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [Guid("F4F6BBBE-2EF4-4E02-B547-1B96271B57E2")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]

    public interface IListString
    {
        // Properties
        int Capacity
        {
            [Description("Gets or sets the total number of elements the internal data structure can hold without resizing.")]
            get;
            [Description("Gets or sets the total number of elements the internal data structure can hold without resizing.")]
            set;
        }

        int Count
        {
            [Description("Gets the number of elements contained in the List<String>.")]
            get;
        }

        [Description("Gets or sets the element at the specified index.")]
        string this[int index]
        {
            get;
            set;
        }

        bool IsReadOnly
        {
            [Description("Gets a value indicating whether the ICollection<T> is read-only.")]
            get;
        }

        bool IsFixedSize
        {
            [Description("Gets a value indicating whether the IList has a fixed size.")]
            get;
        }

        object SyncRoot
        {
            [Description("Gets an object that can be used to synchronize access to the ICollection.")]
            get;
        }

        bool IsSynchronized
        {
            [Description("Gets a value indicating whether access to the ICollection is synchronized (thread safe).")]
            get;
        }

        // Methods

        [Description("Adds an string to the end of the List<string>.")]
        void Add(string value);

        [Description("Adds the elements of the specified collection to the end of the List<string>.")]
        void AddRange(GCollections.IEnumerable collection);

        [Description("Searches the entire sorted List<String> for an element using the default comparer and returns the zero-based index of the element.")]
        int BinarySearch(string item);

        [Description("Searches the entire sorted List<String> for an element using the specified comparer and returns the zero-based index of the element.")]
        int BinarySearch2(string item, IComparer comparer);

        [Description("Searches a range of elements in the sorted List<String> for an element using the specified comparer and returns the zero-based index of the element.")]
        int BinarySearch3(int index, int count, string item, IComparer comparer);

        [Description("Removes all elements from the List<string>.")]
        void Clear();

        [Description("Determines whether an element is in the List<string>.")]
        bool Contains(string value);

        //System.Collections.Generic.List<TOutput> ConvertAll<TOutput>(Converter<T, TOutput> converter);

        [Description("Copies the entire List<string> to a compatible one-dimensional array, starting at the beginning of the target array.")]
        void CopyTo([In][Out] ref string[] array);

        [Description("Copies the entire List<string> to a compatible one-dimensional array, starting at the specified index of the target array.")]
        void CopyTo2([In][Out] ref string[] array, int arrayIndex);

        [Description("Copies a range of elements from the List<string> to a compatible one-dimensional array, starting at the specified index of the target array.")]
        void CopyTo3(int index, [In][Out] ref string[] array, int arrayIndex, int count);

        [Description("Returns an enumerator that iterates through a collection.")]
        GCollections.IEnumerator GetEnumerator();

        [Description("Searches for the specified string and returns the zero-based index of the first occurrence within the entire List<string>.")]
        int IndexOf(string item);

        [Description("Searches for the specified string and returns the zero-based index of the first occurrence within the range of elements in the List<string> that extends from the specified index to the last element.")]
        int IndexOf2(string item, int index);

        [Description("Searches for the specified string and returns the zero-based index of the first occurrence within the range of elements in the List<string> that starts at the specified index and contains the specified number of elements.")]
        int IndexOf3(string item, int index, int count);

        [Description("Inserts an element into the List<T> at the specified index.")]
        void Insert(int index, string item);

        [Description("Searches for the specified string and returns the zero-based index of the last occurrence within the entire List<string>.")]
        int LastIndexOf(string item);

        [Description("Searches for the specified string and returns the zero-based index of the last occurrence within the range of elements in the List<string> that extends from the first element to the specified index.")]
        int LastIndexOf2(string item, int index);

        [Description("Searches for the specified string and returns the zero-based index of the last occurrence within the range of elements in the List<string> that contains the specified number of elements and ends at the specified index.")]
        int LastIndexOf3(string item, int index, int count);

        [Description("Removes the first occurrence of a specific string from the IList.")]
        bool Remove(string value);

        [Description("Removes the element at the specified index of the List<string>.")]
        void RemoveAt(int index);

        [Description("Removes a range of elements from the List<string>.")]
        void RemoveRange(int index, int count);

        [Description("Reverses the order of the elements in the entire List<string>.")]
        void Reverse();

        //public void Sort_2 (Comparison<T> comparison);
        //public void Sort_2 (int index, int count, System.Collections.Generic.IComparer<T> comparer);
        [Description("Sorts the elements in the entire List<string> using the default comparer.")]
        void Sort();

        [Description("Copies the elements of the List<string> to a new array.")]
        string[] ToArray();

        [Description("Sets the capacity to the actual number of elements in the List<string>, if that number is less than a threshold value.")]
        void TrimExcess();

        //[Description("Determines whether every element in the List<T> matches the conditions defined by the specified predicate.")]
        //bool TrueForAll(Predicate<T> match);
    }

}