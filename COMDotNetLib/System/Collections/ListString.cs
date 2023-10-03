using GSystem = global::System; 
using GCollections = global::System.Collections;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a strongly typed list of strings that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [Guid("70EF1DD1-80FA-4276-8051-426C0FAFB2FA")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IListString))]
    public class ListString  : IListString, GCollections.IList, ICollection, IEnumerable
    {
        private GSystem.Collections.Generic.List<string> _stringList;
        public ListString()
        {
            _stringList = new List<string>();
        }

        public ListString(List<String> objListString)
        {
            _stringList = objListString;
        }

        public ListString(int capacity)
        {
            _stringList = new List<string>(capacity);
        }

        public ListString(GSystem.Collections.Generic.IEnumerable<string> collection)
        {
            _stringList = new List<string>(collection);
        }

        // Properties
        public int Capacity
        {
            get { return _stringList.Capacity; }
            set { _stringList.Capacity = value; }
        }

        public int Count
        {
            get { return _stringList.Count; }
        }

        public bool IsReadOnly => ((GCollections.IList)_stringList).IsReadOnly;

        public bool IsFixedSize => ((GCollections.IList)_stringList).IsFixedSize;

        public object SyncRoot => ((ICollection)_stringList).SyncRoot;

        public bool IsSynchronized => ((ICollection)_stringList).IsSynchronized;

        object GCollections.IList.this[int index] { get => ((GCollections.IList)_stringList)[index]; set => ((GCollections.IList)_stringList)[index] = value; }

        public string this[int index]
        {
            get { return _stringList[index]; }
            set { _stringList[index] = value; }
        }

        // Methods
        public void Add(string value)
        {
            _stringList.Add(value);
        }

        //public void AddRange(GSystem.Collections.IEnumerable collection)
        //{
        //    this._stringList.AddRange((IEnumerable<string>)collection);
        //}

        //public int BinarySearch(T item);
        //public int BinarySearch (T item, System.Collections.Generic.IComparer<T> comparer);
        //public int BinarySearch (int index, int count, T item, System.Collections.Generic.IComparer<T> comparer);

        public int BinarySearch(string item)
        {
            return _stringList.BinarySearch(item);
        }

        //public int BinarySearch2(string item, GSystem.Collections.Generic.IComparer<string> comparer)
        //{
        //    return this._stringList.BinarySearch(item, comparer);
        //}

        public void Clear()
        {
            _stringList.Clear();
        }

        public bool Contains(string value)
        {
            return _stringList.Contains(value);
        }

        //public System.Collections.Generic.List<TOutput> ConvertAll<TOutput>(Converter<T, TOutput> converter);

        public void CopyTo([In] [Out] ref string[] array)
        {
            _stringList.CopyTo(array);
        }

        public void CopyTo2([In][Out] ref string[] array, int arrayIndex)
        {
            _stringList.CopyTo(array, arrayIndex);
        }

        public void CopyTo3(int index, [In][Out] ref string[] array, int arrayIndex, int count)
        {
            _stringList.CopyTo(index, array, arrayIndex, count);
        }

        //public bool Exists (Predicate<T> match);

        //public T Find (Predicate<T> match);

        //public System.Collections.Generic.List<T> FindAll (Predicate<T> match);

        //public int FindIndex (int startIndex, int count, Predicate<T> match);
        //public int FindIndex (Predicate<T> match);
        //public int FindIndex (int startIndex, Predicate<T> match);

        //public T FindLast (Predicate<T> match);

        //public int FindLastIndex (Predicate<T> match);
        //public int FindLastIndex (int startIndex, Predicate<T> match);
        //public int FindLastIndex (int startIndex, int count, Predicate<T> match);

        //public void ForEach (Action<T> action);

        //public GSystem.Collections.Generic.List<string>.Enumerator GetEnumerator()
        //{
        //    return this._stringList.GetEnumerator();
        //}

        //public GSystem.Collections.Generic.List<string> GetRange(int index, int count)
        //{
        //    return this._stringList.GetRange(index,count);
        //}

        public int IndexOf(string item)
        {
            return _stringList.IndexOf(item);
        }

        public int IndexOf2(string item, int index)
        {
            return _stringList.IndexOf(item, index);
        }

        public int IndexOf3(string item, int index, int count)
        {
            return _stringList.IndexOf(item, index, count);
        }

        public void Insert(int index, string item)
        {
            _stringList.Insert(index, item);
        }

        //public void InsertRange(int index, System.Collections.Generic.IEnumerable<T> collection);

        public int LastIndexOf(string item)
        {
            return _stringList.LastIndexOf(item);
        }

        public int LastIndexOf2(string item, int index)
        {
            return _stringList.LastIndexOf(item, index);
        }

        public int LastIndexOf3(string item, int index, int count)
        {
            return _stringList.LastIndexOf(item, index, count);
        }

        public bool Remove(string item)
        {
            return _stringList.Remove(item);
        }

        //public int RemoveAll (Predicate<T> match);

        public void RemoveAt(int index)
        {
            _stringList.RemoveAt(index);
        }

        public void RemoveRange(int index, int count)
        {
            _stringList.RemoveRange(index, count);
        }

        public void Reverse()
        {
            _stringList.Reverse();
        }

        //public void Sort (Comparison<string> comparison)
        //{ 
        //    this._stringList.Sort(comparison);
        //}

        //public void Sort (int index, int count, System.Collections.Generic.IComparer<T> comparer);
        public void Sort()
        {
            _stringList.Sort();
        }

        public string[] ToArray()
        { 
            return _stringList.ToArray();
        }

        public void TrimExcess()
        {
            _stringList.TrimExcess();
        }

        public int Add(object value)
        {
            return ((GCollections.IList)_stringList).Add(value);
        }

        public bool Contains(object value)
        {
            return ((GCollections.IList)_stringList).Contains(value);
        }

        public int IndexOf(object value)
        {
            return ((GCollections.IList)_stringList).IndexOf(value);
        }

        public void Insert(int index, object value)
        {
            ((GCollections.IList)_stringList).Insert(index, value);
        }

        public void Remove(object value)
        {
            ((GCollections.IList)_stringList).Remove(value);
        }

        public void CopyTo(Array array, int index)
        {
            ((ICollection)_stringList).CopyTo(array, index);
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)_stringList).GetEnumerator();
        }

        //public bool TrueForAll(Predicate<T> match);

    }
}
