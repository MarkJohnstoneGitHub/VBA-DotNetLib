// https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist?view=netframework-4.8.1

using GSystem = global::System;
using GCollections = global::System.Collections;
using System.Collections;
using DotNetLib.Extensions;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a collection of key/value pairs that are sorted by the keys and are accessible by key and by index.")]
    [Guid("65009713-67CA-4DF1-A29C-B37C64AB8B5D")]
    [ProgId("DotNetLib.System.Collections.SortedList")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ISortedList))]
    public class SortedList : GCollections.IDictionary, GCollections.ICollection, GCollections.IEnumerable, GSystem.ICloneable, IDictionary, ICollection, IWrappedObject, ISortedList
    {
        private GCollections.SortedList _sortedList;

        // Constructors

        public SortedList()
        {
            _sortedList = new GCollections.SortedList();
        }

        public SortedList(int initialCapacity)
        {
            _sortedList = new GCollections.SortedList(initialCapacity);
        }

        public SortedList(GCollections.IComparer comparer)
        {
            _sortedList = new GCollections.SortedList(comparer);
        }

        public SortedList(GCollections.IComparer comparer, int capacity)
        {
            _sortedList = new GCollections.SortedList(comparer, capacity);
        }

        public SortedList(GCollections.IDictionary d)
        {
            _sortedList = new GCollections.SortedList(d);
        }

        public SortedList(GCollections.IDictionary d, GCollections.IComparer comparer)
        {
            _sortedList = new GCollections.SortedList(d, comparer);
        }

        // Properties
        public object WrappedObject => _sortedList;

        internal GCollections.SortedList WrappedSortedList
        {
            get { return _sortedList; }
        }

        public int Capacity 
        {
            get => _sortedList.Capacity;
            set => _sortedList.Capacity = value;
        }

        public int Count => _sortedList.Count;

        public bool IsReadOnly => _sortedList.IsReadOnly;

        public bool IsFixedSize => _sortedList.IsFixedSize;

        public bool IsSynchronized => _sortedList.IsSynchronized;

        public object this[object key]
        {
            get => _sortedList[key];
            set => _sortedList[key] = value;
        }

        public ICollection Keys => (ICollection)_sortedList.Keys;

        public object SyncRoot => _sortedList.SyncRoot;

        public ICollection Values => (ICollection)_sortedList.Values;

        GCollections.ICollection GCollections.IDictionary.Keys => _sortedList.Keys;

        GCollections.ICollection GCollections.IDictionary.Values => _sortedList.Values;


        // Methods

        public void Add(object key, object value)
        {
            _sortedList.Add(key, value);
        }

        public void Clear()
        {
            _sortedList.Clear();
        }

        public object Clone()
        {
            return _sortedList.Clone();
        }

        public bool Contains(object key)
        {
            return _sortedList.Contains(key);
        }

        public bool ContainsKey(object key)
        {
            return _sortedList.ContainsKey(key);
        }

        public bool ContainsValue(object value)
        {
            return _sortedList.ContainsValue(value);
        }

        public void CopyTo(global::System.Array array, int index)
        {
            _sortedList.CopyTo(array, index);
        }

        public void CopyTo(Array array, int arrayIndex)
        {
            _sortedList.CopyTo(array.WrappedArray, arrayIndex);
        }

        public object GetByIndex(int index)
        {
            return _sortedList.GetByIndex(index);
        }

        public IDictionaryEnumerator GetEnumerator()
        {
            return _sortedList.GetEnumerator();
        }

        public object GetKey(int index)
        {
            return _sortedList.GetKey(index);
        }

        public GCollections.IList GetKeyList()
        {
            return _sortedList.GetKeyList();
        }

        public GCollections.IList GetValueList()
        {
            return _sortedList.GetValueList();
        }

        public int IndexOfKey(object key)
        {
            return _sortedList.IndexOfKey(key);
        }

        public int IndexOfValue(object value)
        {
            return _sortedList.IndexOfKey(value);
        }

        public void Remove(object key)
        {
            _sortedList.Remove(key);
        }

        public void RemoveAt(int index)
        {
            _sortedList.Remove(index);
        }

        public void SetByIndex(int index, object value)
        {
            _sortedList.SetByIndex(index, value);
        }

        public void TrimToSize()
        {
            _sortedList.TrimToSize();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _sortedList.GetEnumerator();
        }

        public new Type GetType()
        {
            return new Type(((object)this).GetType());
        }
    }
}
