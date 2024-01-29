// https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1?view=netframework-4.8.1

using GCollections = global::System.Collections;

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("BF3921C9-7281-4C35-A689-E60375A54ED5")]
    [Description("Represents a list of objects that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IListObject
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
            [Description("Gets the number of elements contained in the List<Object>.")]
            get;
        }

        object this[int index]
        {
            [Description("Gets or sets the element at the specified index.")]
            get;
            [Description("Gets or sets the element at the specified index.")]
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

        // Methods

        [Description("Adds an object to the end of the List<T>.")]
        void Add(object value);

        [Description("Adds the elements of the specified collection to the end of the List<T>.")]
        void AddRange(GCollections.IEnumerable collection);

        [Description("Searches the entire sorted List<object> for an element using the default comparer and returns the zero-based index of the element.")]
        int BinarySearch(object item);

        [Description("Searches the entire sorted List<object> for an element using the specified comparer and returns the zero-based index of the element.")]
        int BinarySearch2(object item, IComparer comparer);

        [Description("Searches a range of elements in the sorted List<object> for an element using the specified comparer and returns the zero-based index of the element.")]
        int BinarySearch3(int index, int count, object item, IComparer comparer);

        [Description("Removes all elements from the List<T>.")]
        void Clear();

        [Description("Determines whether an element is in the List<T>.")]
        bool Contains(object value);

        [Description("Copies the entire List<T> to a compatible one-dimensional array, starting at the beginning of the target array.")]
        void CopyTo([In][Out] ref object[] array);

        [Description("Copies the entire List<T> to a compatible one-dimensional array, starting at the specified index of the target array.")]
        void CopyTo2([In][Out] ref object[] array, int arrayIndex);

        [Description("Copies a range of elements from the List<T> to a compatible one-dimensional array, starting at the specified index of the target array.")]
        void CopyTo3(int index, [In][Out] ref object[] array, int arrayIndex, int count);

        [Description("Returns an enumerator that iterates through a collection.")]
        GCollections.IEnumerator GetEnumerator();

        [Description("Searches for the specified object and returns the zero-based index of the first occurrence within the entire List<string>.")]
        int IndexOf(object item);

        [Description("Searches for the specified object and returns the zero-based index of the first occurrence within the range of elements in the List<string> that extends from the specified index to the last element.")]
        int IndexOf2(object item, int index);

        [Description("Searches for the specified object and returns the zero-based index of the first occurrence within the range of elements in the List<string> that starts at the specified index and contains the specified number of elements.")]
        int IndexOf3(object item, int index, int count);

        [Description("Inserts an element into the List<T> at the specified index.")]
        void Insert(int index, object item);

        [Description("Searches for the specified List<object> and returns the zero-based index of the last occurrence within the entire List<object>.")]
        int LastIndexOf(object item);

        [Description("Searches for the specified object and returns the zero-based index of the last occurrence within the range of elements in the List<string> that extends from the first element to the specified index.")]
        int LastIndexOf2(object item, int index);

        [Description("Searches for the specified object and returns the zero-based index of the last occurrence within the range of elements in the List<string> that contains the specified number of elements and ends at the specified index.")]
        int LastIndexOf3(object item, int index, int count);

        [Description("Removes the first occurrence of a specific object from the IList.")]
        bool Remove(object value);

        [Description("Removes the element at the specified index of the List<object>.")]
        void RemoveAt(int index);

        [Description("Removes a range of elements from the List<object>.")]
        void RemoveRange(int index, int count);

        [Description("Reverses the order of the elements in the entire List<object>.")]
        void Reverse();

        //public void Sort2 (Comparison<T> comparison);
        //public void Sort2 (int index, int count, System.Collections.Generic.IComparer<T> comparer);
        [Description("Sorts the elements in the entire List<object> using the default comparer.")]
        void Sort();

        [Description("Copies the elements of the List<object> to a new array.")]
        object[] ToArray();

        [Description("Sets the capacity to the actual number of elements in the List<object>, if that number is less than a threshold value.")]
        void TrimExcess();

        //[Description("Determines whether every element in the List<T> matches the conditions defined by the specified predicate.")]
        //bool TrueForAll(Predicate<T> match);


    }
}
