// https://learn.microsoft.com/en-us/dotnet/api/system.array?view=netframework-4.8.1

using GArray = global::System.Array;
using GSystem = global::System;
using System;
using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Description("Provides methods for creating, manipulating, searching, and sorting arrays, thereby serving as the base class for all arrays in the common language runtime.")]
    [Guid("52F63C79-278B-4B47-A4C3-C50126CC2E76")]
    [ProgId("DotNetLib.System.ArraySingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IArraySingleton))]
    public class ArraySingleton : IArraySingleton
    {
        public ArraySingleton() { }

        //public GLinq.IQueryable AsQueryable(this IEnumerable source)
        //{
        //    throw new NotImplementedException();
        //}

        // Note added to wrap a mscorlib.Array for mscorlib.ICollection void CopyTo2(Array array, int index);
        // If array null throw error??
        public Array Create(GSystem.Array array)
        { 
            return new Array(array);
        }

        public int BinarySearch(Array array, object value)
        {
            return Array.BinarySearch(array, value);
        }

        public int BinarySearch2(Array array, object value, IComparer comparer = null)
        {
            return Array.BinarySearch(array, value, comparer);
        }

        public int BinarySearch3(Array array, int index, int length, object value, IComparer comparer = null)
        {
            return Array.BinarySearch(array, index, length, value, comparer);
        }

        public void Clear(Array array, int index, int length)
        {
            Array.Clear(array, index, length);
        }

        public void ConstrainedCopy(Array sourceArray, int sourceIndex, Array destinationArray, int destinationIndex, int length)
        {
            throw new NotImplementedException();
        }

        public void Copy(Array sourceArray, Array destinationArray, int length)
        {
            Array.Copy(sourceArray, destinationArray, length);
        }

        public void Copy2(Array sourceArray, Array destinationArray, long length)
        {
            Array.Copy(sourceArray, destinationArray, length);
        }

        public void Copy3(Array sourceArray, int sourceIndex, Array destinationArray, int destinationIndex, int length)
        {
            Array.Copy(sourceArray,sourceIndex, destinationArray, destinationIndex, length);
        }

        public Array CreateInstance(Type elementType, int length)
        {
            return Array.CreateInstance(elementType, length);
        }

        public Array CreateInstance2(Type elementType, int length1, int length2)
        {
            return Array.CreateInstance(elementType, length1, length2);
        }
        public Array CreateInstance3(Type elementType, int length1, int length2, int length3)
        {
            return Array.CreateInstance(elementType, length1, length2, length3);
        }

        public Array CreateInstance4(Type elementType, [In] ref int[] lengths)
        {
            return Array.CreateInstance(elementType, lengths);
        }

        public Array CreateInstance5(Type elementType, [In] ref long[] lengths)
        {
            return Array.CreateInstance(elementType, lengths);
        }

        public Array CreateInstance6(Type elementType, [In] ref int[] lengths, [In] ref int[] lowerBounds)
        {
            return Array.CreateInstance(elementType, lengths,lowerBounds);
        }

        public object Find(Array array, Predicate match)
        {
            return Array.Find<object>(array, match);
        }

       public Array FindAll(Array array, Predicate match)
       {
            return Array.FindAll<object>(array, match);
       }

        public int FindIndex(Array array, Predicate match)
        {
            return Array.FindIndex<object>(array, match);
        }

        public int FindIndex2(Array array, int startIndex, Predicate match)
        {
            return Array.FindIndex<object>(array, startIndex, match);
        }

        public int FindIndex3(Array array, int startIndex, int count, Predicate match)
        {
            return Array.FindIndex<object>(array, startIndex, count, match);
        }

        public object FindLast(Array array, Predicate match)
        {
            return Array.FindLast<object>(array, match);
        }

        public int FindLastIndex(Array array, Predicate match)
        {
            return Array.FindLastIndex<object>(array, match);
        }

        public int FindLastIndex2(Array array, int startIndex, Predicate match)
        {
            return Array.FindLastIndex<object>(array, startIndex, match);
        }

        public int FindLastIndex3(Array array, int startIndex, int count, Predicate match)
        {
            return Array.FindLastIndex<object>(array, startIndex, count, match);
        }

        public int IndexOf(Array array, object value)
        {
            return Array.IndexOf(array,value);
        }

        public int IndexOf2(Array array, object value, int startIndex)
        {
            return Array.IndexOf(array, value, startIndex);
        }

        public int IndexOf3(Array array, object value, int startIndex, int count)
        {
            return Array.IndexOf(array, value, startIndex, count);
        }

        public int LastIndexOf(Array array, object value)
        {
            return Array.LastIndexOf(array, value);
        }

        public int LastIndexOf2(Array array, object value, int startIndex)
        {
            return Array.LastIndexOf(array, value, startIndex);
        }

        public int LastIndexOf3(Array array, object value, int startIndex, int count)
        {
            return Array.LastIndexOf(array, value, startIndex, count);
        }

        public void Resize(ref Array array, int newSize)
        {
            Array.Resize(ref array, newSize);
        }

        public void Reverse(Array array)
        {
            Array.Reverse(array);
        }

        public void Reverse2(Array array, int index, int length)
        {
            Array.Reverse(array, index, length);
        }

        //public void Sort2(Array array)
        //{
        //    Array.Sort2(array);
        //}

        public void Sort(Array array, IComparer comparer = null)
        {
            Array.Sort(array, comparer);
        }

        //public void Sort2(Array array, int index, int length)
        //{
        //    Array.Sort2(array, index, length);
        //}

        public void Sort2(Array array, int index, int length, IComparer comparer = null)
        {
            Array.Sort(array, index, length, comparer);
        }

        //public void Sort2(Array keys, Array items)
        //{
        //    Array.Sort2(keys, items);
        //}

        public void Sort3(Array keys, Array items, IComparer comparer = null)
        {
            Array.Sort(keys, items, comparer);
        }

        //public void Sort2(Array keys, Array items, int index, int length)
        //{
        //    Array.Sort2(keys, items, index, length);
        //}

        public void Sort4(Array keys, Array items, int index, int length, IComparer comparer = null)
        {
            Array.Sort(keys, items, index, length, comparer);
        }

        public bool TrueForAll(Array array, Predicate match)
        {
            return Array.TrueForAll<object>(array, match);
        }
    }
}


//public int BinarySearch2(Array array, object value)
//{
//    return Array.BinarySearch2(array, value);
//}

//public int BinarySearch2(Array array, int index, int length, object value)
//{
//    throw new NotImplementedException();
//}