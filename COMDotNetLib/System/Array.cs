// https://learn.microsoft.com/en-us/dotnet/api/system.array?view=netframework-4.8.1

using GSystem = global::System;
using GArray = global::System.Array;
using GCollections = global::System.Collections;
using System;
using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.Extensions;
using System.Diagnostics.Contracts;
using System.Runtime.ConstrainedExecution;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Description("Provides methods for creating, manipulating, searching, and sorting arrays, thereby serving as the base class for all arrays in the common language runtime.")]
    [Guid("DC1795C5-9961-438E-9878-9A191B507B5C")]
    [ProgId("DotNetLib.System.Array")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IArray))]
    public class Array : IArray, ICloneable, IList, ICollection, IEnumerable, IStructuralComparable, IStructuralEquatable
    {
        private GSystem.Array _array;

        // Constructors

        public Array(GSystem.Array array)
        {
            _array = array;
        }

        //Properties

        public GSystem.Array WrappedArray
        {
            get { return _array; }
            set { _array = value; }
        }

        public virtual bool IsFixedSize => _array.IsFixedSize;

        public bool IsReadOnly => _array.IsReadOnly;

        public bool IsSynchronized => _array.IsSynchronized;

        public int Length => _array.Length;

        public long LongLength => _array.LongLength;

        public int Rank => _array.Rank;

        public object SyncRoot => _array.SyncRoot;


        // Methods

        public static int BinarySearch(Array array, object value)
        {
            return GArray.BinarySearch(array.WrappedArray, value);
        }

        public static int BinarySearch(Array array, object value, GCollections.IComparer comparer)
        {
            return GArray.BinarySearch(array.WrappedArray,value,comparer);
        }

        public static int BinarySearch(Array array, int index, int length, object value)
        {
            return GArray.BinarySearch(array.WrappedArray, index, length, value);
        }

        public static int BinarySearch(Array array, int index, int length, object value, GCollections.IComparer comparer)
        {
            return GArray.BinarySearch(array.WrappedArray, index, length, value, comparer);
        }

        public static void Clear(Array array, int index, int length)
        {
            GArray.Clear(array.WrappedArray, index, length);
        }

        public object Clone()
        {
            return new Array((GSystem.Array)WrappedArray.Clone());
        }

        public static void ConstrainedCopy(Array sourceArray, int sourceIndex, Array destinationArray, int destinationIndex, int length)
        {
            GArray.ConstrainedCopy(sourceArray.WrappedArray, sourceIndex, destinationArray.WrappedArray, destinationIndex, length);
        }

        public static void Copy(Array sourceArray, Array destinationArray, int length)
        {
            GArray.Copy(sourceArray.WrappedArray, destinationArray.WrappedArray, length);
        }

        public static void Copy(Array sourceArray, Array destinationArray, long length)
        {
            GArray.Copy(sourceArray.WrappedArray, destinationArray.WrappedArray, length);
        }

        public static void Copy(Array sourceArray, int sourceIndex, Array destinationArray, int destinationIndex, int length)
        {
            GArray.Copy(sourceArray.WrappedArray, sourceIndex, destinationArray.WrappedArray, destinationIndex, length);
        }

        public void CopyTo(Array array, int index)
        {
            _array.CopyTo(array.WrappedArray, index);
        }

        public void CopyTo(Array array, long index)
        {
            _array.CopyTo(array.WrappedArray, index);
        }

        public static Array CreateInstance(Type elementType, int length)
        {
            return new Array(GArray.CreateInstance(elementType.WrappedType, length));
        }

        public static Array CreateInstance(Type elementType, int[] lengths)
        {
            return new Array(GArray.CreateInstance(elementType.WrappedType, lengths));
        }

        public static Array CreateInstance(Type elementType, long[] lengths)
        {
            return new Array(GArray.CreateInstance(elementType.WrappedType, lengths));
        }

        public static Array CreateInstance(Type elementType, int length1, int length2)
        {
            return new Array(GArray.CreateInstance(elementType.WrappedType,length1,length2));
        }

        public static Array CreateInstance(Type elementType, int[] lengths, int[] lowerBounds)
        {
            return new Array(GArray.CreateInstance(elementType.WrappedType, lengths,lowerBounds));
        }

        public static Array CreateInstance(Type elementType, int length1, int length2, int length3)
        {
            return new Array(GArray.CreateInstance(elementType.WrappedType, length1, length2,length3));
        }

        public new virtual bool Equals(object obj)
        {
            return _array.Equals(obj.Unwrap());
        }

        public IEnumerator GetEnumerator()
        {
            return _array.GetEnumerator();
        }

        public new virtual int GetHashCode()
        { 
            return _array.GetHashCode(); 
        }

        public int GetLength(int dimension) 
        { 
            return _array.GetLength(dimension);
        }

        public long GetLongLength(int dimension)
        {
            return _array.GetLongLength(dimension);
        }

        public int GetLowerBound(int dimension)
        {
            return _array.GetLowerBound(dimension);
        }

        public new Type GetType()
        {
            return new Type(_array.GetType());
        }

        public int GetUpperBound(int dimension)
        { 
            return _array.GetUpperBound(dimension);
        }

        public object GetValue(int index)
        {
            return _array.GetValue(index);
        }

        public object GetValue(int index1, int index2)
        {
            return _array.GetValue(index1, index2);
        }

        public object GetValue(int index1, int index2, int index3)
        {
            return _array.GetValue(index1, index2, index3);
        }

        public object GetValue([In] ref int[] indices)
        { 
            return _array.GetValue(indices);
        }

        public object GetValue(long index)
        { 
            return _array.GetValue(index);
        }

        public object GetValue(long index1, long index2)
        {
            return _array.GetValue(index1, index2);
        }

        public object GetValue(long index1, long index2, long index3)
        {
            return _array.GetValue(index1, index2, index3);
        }

        public object GetValue([In] ref long[] indices)
        { 
            return _array.GetValue(indices); 
        }

        public static int IndexOf(Array array, object value)
        { 
            return GArray.IndexOf(array.WrappedArray,value); 
        }

        public static int IndexOf(Array array, object value, int startIndex)
        {
            return GArray.IndexOf(array.WrappedArray, value, startIndex);
        }

        public static int IndexOf(Array array, object value, int startIndex, int count)
        {
            return GArray.IndexOf(array.WrappedArray, value, startIndex, count);
        }

        public void Initialize()
        {  
            _array.Initialize();
        }

        public static int LastIndexOf(Array array, object value)
        {
            return GArray.LastIndexOf(array.WrappedArray, value);
        }

        public static int LastIndexOf(Array array, object value, int startIndex)
        {
            return GArray.LastIndexOf(array.WrappedArray, value, startIndex);
        }

        public static int LastIndexOf(Array array, object value, int startIndex, int count)
        {
            return GArray.LastIndexOf(array.WrappedArray, value, startIndex, count);
        }

        public static void Resize(ref Array array, int newSize)
        {
            GArray newArray = array.WrappedArray;
            Resize(ref newArray, newSize);
            array = new Array(newArray);
        }

        // Converted generic Resize to use non-generic method
        // https://referencesource.microsoft.com/#mscorlib/system/array.cs,50
        // https://stackoverflow.com/a/2085186/10759363
        [ReliabilityContract(Consistency.WillNotCorruptState, Cer.MayFail)]
        internal static void Resize(ref GArray array, int newSize)
        {
            if (newSize < 0)
                throw new ArgumentOutOfRangeException("newSize", "Index is less than zero.");
            Contract.Ensures(Contract.ValueAtReturn(out array) != null);
            Contract.Ensures(Contract.ValueAtReturn(out array).Length == newSize);
            Contract.EndContractBlock();

            GArray larray = array;
            if (larray == null)
            {
                array = GArray.CreateInstance(array.GetType().GetElementType(), newSize);
            }
            else if (larray.Length != newSize)
            {
                GArray newArray = GArray.CreateInstance(array.GetType().GetElementType(), newSize);
                GArray.Copy(larray, 0, newArray, 0, (larray.Length > newSize) ? newSize : larray.Length);
                array = newArray;
            }
        }

        // https://referencesource.microsoft.com/#mscorlib/system/array.cs,50
        //[ReliabilityContract(Consistency.WillNotCorruptState, Cer.MayFail)]
        //public static void Resize<T>(ref T[] array, int newSize)
        //{
        //    if (newSize < 0)
        //        throw new ArgumentOutOfRangeException("newSize", Environment.GetResourceString("ArgumentOutOfRange_NeedNonNegNum"));
        //    Contract.Ensures(Contract.ValueAtReturn(out array) != null);
        //    Contract.Ensures(Contract.ValueAtReturn(out array).Length == newSize);
        //    Contract.EndContractBlock();

        //    T[] larray = array;
        //    if (larray == null)
        //    {
        //        array = new T[newSize];
        //        return;
        //    }

        //    if (larray.Length != newSize)
        //    {
        //        T[] newArray = new T[newSize];
        //        Array.Copy(larray, 0, newArray, 0, larray.Length > newSize ? newSize : larray.Length);
        //        array = newArray;
        //    }
        //}

        public static void Reverse(Array array)
        {
            GArray.Reverse(array.WrappedArray);
        }

        public static void Reverse(Array array, int index, int length)
        {
            GArray.Reverse(array.WrappedArray, index, length);
        }

        public void SetValue(object value, int index)
        {
            _array.SetValue(value, index);
        }

        public void SetValue(object value, int index1, int index2)
        {
            _array.SetValue(value, index1,  index2);
        }

        public void SetValue(object value, int index1, int index2, int index3)
        {
            _array.SetValue(value, index1, index2, index3);
        }

        public void SetValue(object value, [In] ref int[] indices)
        {
            _array.SetValue(value, indices);
        }

        public void SetValue(object value, long index)
        {
            _array.SetValue(value, index);
        }

        public void SetValue(object value, long index1, long index2)
        {
            _array.SetValue(value, index1, index2);
        }

        public void SetValue(object value, long index1, long index2, long index3)
        {
            _array.SetValue(value, index1, index2, index3);
        }

        public void SetValue(object value, [In] ref long[] indices)
        {
            _array.SetValue(value, indices);
        }

        public static void Sort(Array array)
        {
            GArray.Sort(array.WrappedArray);
        }

        public static void Sort(Array array, IComparer comparer)
        {
            GArray.Sort(array.WrappedArray, comparer);
        }

        public static void Sort(Array array, int index, int length)
        {
            GArray.Sort(array.WrappedArray, index, length);
        }

        public static void Sort(Array array, int index, int length, IComparer comparer)
        {
            GArray.Sort(array.WrappedArray, index, length, comparer);
        }

        public static void Sort(Array keys, Array items)
        {
            GArray.Sort(keys.WrappedArray, items.WrappedArray);
        }

        public static void Sort(Array keys, Array items, IComparer comparer)
        {
            GArray.Sort(keys.WrappedArray, items.WrappedArray, comparer);
        }

        public static void Sort(Array keys, Array items, int index, int length)
        {
            GArray.Sort(keys.WrappedArray, items.WrappedArray, index, length);
        }

        public static void Sort(Array keys, Array items, int index, int length, IComparer comparer)
        {
            GArray.Sort(keys.WrappedArray, items.WrappedArray, index, length, comparer);
        }


        // Explicit Interface Implementations
        // https://learn.microsoft.com/en-us/dotnet/api/system.array?view=netframework-4.8.1#explicit-interface-implementations

        public int Count => ((ICollection)_array).Count;

        public object this[int index]
        {
            get => ((IList)_array)[index];
            set => ((IList)_array)[index] = value;
        }

        public int Add(object value)
        {
            throw new NotSupportedException();
            //    return ((IList)_array).Add(value);
        }

        public void Clear()
        {
            ((IList)_array).Clear();
        }

        public bool Contains(object value)
        {
            return ((IList)_array).Contains(value);
        }

        public void CopyTo(GSystem.Array array, int index)
        {
            ((ICollection)_array).CopyTo(array, index);
        }

        public int IndexOf(object value)
        {
            return ((IList)_array).IndexOf(value);
        }

        public void Insert(int index, object value)
        {
            ((IList)_array).Insert(index, value);
        }

        public void Remove(object value)
        {
            ((IList)_array).Remove(value);
        }

        public void RemoveAt(int index)
        {
            ((IList)_array).RemoveAt(index);
        }

        public int CompareTo(object other, IComparer comparer)
        {
            return ((IStructuralComparable)_array).CompareTo(other, comparer);
        }

        public bool Equals(object other, IEqualityComparer comparer)
        {
            return ((IStructuralEquatable)_array).Equals(other, comparer);
        }

        public int GetHashCode(IEqualityComparer comparer)
        {
            return ((IStructuralEquatable)_array).GetHashCode(comparer);
        }

        new public virtual string ToString()
        {
            return _array.ToString();

        }

    }
}
