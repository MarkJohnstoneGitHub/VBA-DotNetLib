// https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1?view=netframework-4.8.1

using GGeneric = global::System.Collections.Generic;
using GCollections = global::System.Collections;
using System;
using System.Collections.Generic;
using System.Collections;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a objects that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [Guid("B8EFD58F-2A9F-47DE-9473-C322520A01CD")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IListObject))]
    public class ListObject : GCollections.IList, GCollections.ICollection, IEnumerable, IListObject
    {
        private GGeneric.List<object> _list;

        public ListObject()
        {
            _list = new List<object>();
        }

        public ListObject(int capacity)
        {
            _list = new List<object>(capacity);
        }

        public ListObject(GCollections.IEnumerable collection)
        {
            _list = new List<object>((IEnumerable<object>)collection);
        }

        // Properties
        public int Capacity
        {
            get { return _list.Capacity; }
            set { _list.Capacity = value; }
        }

        public int Count
        {
            get { return _list.Count; }
        }

        public bool IsReadOnly => ((GCollections.IList)_list).IsReadOnly;

        public bool IsFixedSize => ((GCollections.IList)_list).IsFixedSize;

        public object SyncRoot => ((GCollections.ICollection)_list).SyncRoot;

        public bool IsSynchronized => ((GCollections.ICollection)_list).IsSynchronized;

        //object GCollections.IList.this[int index] { get => ((GCollections.IList)_list)[index]; set => ((GCollections.IList)_list)[index] = value; }

        public object this[int index]
        {
            get { return _list[index]; }
            set { _list[index] = value; }
        }

        // Methods
        public void Add(object value)
        {
            _list.Add(value);
        }

        public void AddRange(IEnumerable collection)
        {
            _list.AddRange((IEnumerable<object>)collection);
        }

        public int BinarySearch(object item)
        {
            return _list.BinarySearch(item);
        }

        public int BinarySearch2(object item, IComparer comparer)
        {
            return _list.BinarySearch(item, (IComparer<object>)comparer);
        }

        public int BinarySearch3(int index, int count, object item, IComparer comparer)
        {
            return _list.BinarySearch(index, count, item, (IComparer<object>)comparer);
        }

        public void Clear()
        {
            _list.Clear();
        }

        public bool Contains(object value)
        {
            return _list.Contains(value);
        }

        public void CopyTo([In][Out] ref object[] array)
        {
            _list.CopyTo(array);
        }

        public void CopyTo2([In][Out] ref object[] array, int arrayIndex)
        {
            _list.CopyTo(array, arrayIndex);
        }

        public void CopyTo3(int index, [In][Out] ref object[] array, int arrayIndex, int count)
        {
            _list.CopyTo(index, array, arrayIndex, count);
        }

        public int IndexOf(object item)
        {
            return _list.IndexOf(item);
        }

        public int IndexOf2(object item, int index)
        {
            return _list.IndexOf(item, index);
        }

        public int IndexOf3(object item, int index, int count)
        {
            return _list.IndexOf(item, index, count);
        }

        public void Insert(int index, object item)
        {
            _list.Insert(index, item);
        }

        public int LastIndexOf(object item)
        {
            return _list.LastIndexOf(item);
        }

        public int LastIndexOf2(object item, int index)
        {
            return _list.LastIndexOf(item, index);
        }

        public int LastIndexOf3(object item, int index, int count)
        {
            return _list.LastIndexOf(item, index, count);
        }

        public bool Remove(object value)
        {
            return _list.Remove(value);
        }

        //public int RemoveAll (Predicate<T> match);

        public void RemoveAt(int index)
        {
            _list.RemoveAt(index);
        }

        public void RemoveRange(int index, int count)
        {
            _list.RemoveRange(index, count);
        }

        public void Reverse()
        {
            _list.Reverse();
        }

        //public void Sort2 (Comparison<string> comparison)
        //{ 
        //    this._stringList.Sort2(comparison);
        //}

        //public void Sort2 (int index, int count, System.Collections.Generic.IComparer<T> comparer);
        public void Sort()
        {
            _list.Sort();
        }

        public object[] ToArray()
        {
            return _list.ToArray();
        }

        public void TrimExcess()
        {
            _list.TrimExcess();
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)_list).GetEnumerator();
        }

        int GCollections.IList.Add(object value)
        {
            return ((GCollections.IList)_list).Add(value);
        }

        void GCollections.IList.Remove(object value)
        {
            ((GCollections.IList)_list).Remove(value);
        }

        public void CopyTo(Array array, int index)
        {
            ((GCollections.ICollection)_list).CopyTo(array.WrappedArray, index);
        }

        public void CopyTo(global::System.Array array, int index)
        {
            ((GCollections.ICollection)_list).CopyTo(array, index);
        }
    }
}
