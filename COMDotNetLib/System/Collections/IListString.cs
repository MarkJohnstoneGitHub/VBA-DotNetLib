using System.ComponentModel;
using System.Runtime.InteropServices;
using System;
using GSystem = global::System;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a strongly typed list of strings that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [Guid("F4F6BBBE-2EF4-4E02-B547-1B96271B57E2")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]

    public interface IListString
    {
        // Constructors
        [Description("Initializes a new instance of the List<string> class that is empty and has the default initial capacity.")]
        ListString Create();

        [Description("Initializes a new instance of the List<string> class that is empty and has the specified initial capacity.")]
        ListString Create2(int capacity);

        [Description("Initializes a new instance of the List<string> class that contains elements copied from the specified collection and has sufficient capacity to accommodate the number of elements copied.")]
        ListString CreateFromIEnumerable(GSystem.Collections.IEnumerable collection);


        // Properties

        [Description("Gets or sets the element at the specified index.")]
        string this[int index]
        {
            get;
            set;
        }

        // Methods

        [Description("Adds an string to the end of the List<string>.")]
        void Add(string value);

        //[Description("Adds the elements of the specified collection to the end of the List<string>.")]
        //void AddRange(GSystem.Collections.IEnumerable collection);

        [Description("Removes all elements from the List<string>.")]
        void Clear();

        [Description("Determines whether an element is in the List<string>.")]
        bool Contains(string value);

        //System.Collections.Generic.List<TOutput> ConvertAll<TOutput>(Converter<T, TOutput> converter);

        [Description("Copies the entire List<string> to a compatible one-dimensional array, starting at the beginning of the target array.")]
        void CopyTo(string[] array);

        [Description("Copies the entire List<string> to a compatible one-dimensional array, starting at the specified index of the target array.")]
        void CopyTo2(string[] array, int arrayIndex);

        [Description("Copies a range of elements from the List<string> to a compatible one-dimensional array, starting at the specified index of the target array.")]
        void CopyTo3(int index, string[] array, int arrayIndex, int count);

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

        //public void Sort (Comparison<T> comparison);
        //public void Sort (int index, int count, System.Collections.Generic.IComparer<T> comparer);
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