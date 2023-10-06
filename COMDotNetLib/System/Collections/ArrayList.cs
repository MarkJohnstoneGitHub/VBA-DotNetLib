// https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Implements the IList interface using an array whose size is dynamically increased as required.")]
    [Guid("CE76B55C-7BDD-4471-90B8-704298D3BDF3")]
    [ProgId("DotNetLib.System.Collections.ArrayList")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IArrayList))]
    public class ArrayList : IArrayList,  ICloneable, GCollections.IList
    {
        private GCollections.ArrayList _arrayList;

        // Constructors

        public ArrayList()
        {
            _arrayList = new GCollections.ArrayList();
        }

        public ArrayList(GCollections.ArrayList arrayList)
        {
            _arrayList = arrayList;
        }

        public ArrayList(int capacity)
        {
            _arrayList = new GCollections.ArrayList(capacity);
        }

        public ArrayList(GCollections.ICollection c)
        {
            _arrayList = new GCollections.ArrayList(c);
        }

        public ArrayList(GCollections.IList list)
        {
            _arrayList = new GCollections.ArrayList(list); 
        }

        //Properties

        internal GCollections.ArrayList WrappedArrayList
        {
            get { return _arrayList; }
            set { _arrayList = value; }  
        }

        public virtual int Capacity
        {
            get { return _arrayList.Capacity; }
            set { _arrayList.Capacity = value; }

        }

        public virtual int Count => _arrayList.Count;

        public virtual bool IsFixedSize => _arrayList.IsFixedSize;

        public bool IsReadOnly => _arrayList.IsReadOnly;

        public object SyncRoot => _arrayList.SyncRoot;

        public bool IsSynchronized => _arrayList.IsSynchronized;

        // Todo Issue assigning value types added member SetItem(index,item) 
        // https://stackoverflow.com/questions/9481140/exposing-property-as-variant-in-net-for-interop
        // https://stackoverflow.com/a/9924325/10759363
        // https://social.msdn.microsoft.com/Forums/en-US/b8e26285-1f2a-4a1a-9ca4-9d198d0bd9dd/com-interop-property-getletset-interface-attribute?forum=vblanguage
        public object this[int index]
        {
            get => _arrayList[index];
            set => _arrayList[index] = value;
        }

        // Methods

        // Added to fix issue assigning value types.
        public void SetItem(int index, object value)
        {
            _arrayList[index] = value;
        }

        public static ArrayList Adapter(GCollections.IList list)
        { 
            return new ArrayList(GCollections.ArrayList.Adapter(list)); 
        }

        public virtual int Add(object value)
        {
            return _arrayList.Add(value);
        }

        public virtual void AddRange(GCollections.IEnumerable collection)
        {
            _arrayList.Add(collection); 
        }

        public virtual int BinarySearch(object item)
        {
            return _arrayList.BinarySearch(item);
        }

        public virtual int BinarySearch2(object item, GCollections.IComparer comparer)
        {
            return _arrayList.BinarySearch(item,comparer);
        }

        public virtual int BinarySearch3(int index, int count, object item, GCollections.IComparer comparer)
        {
            return _arrayList.BinarySearch(index,count, item, comparer);
        }

        public virtual void Clear()
        {
            _arrayList.Clear();
        }

        public virtual object Clone()
        {
            return new ArrayList((GCollections.ArrayList)_arrayList.Clone());
        }

        public virtual bool Contains(object value)
        {
            return _arrayList.Contains(value);
        }

        public virtual void CopyTo(Array array)
        {
            _arrayList.CopyTo(array);
        }

        public virtual void CopyTo2(Array array, int arrayIndex)
        {
            _arrayList.CopyTo(array,arrayIndex);
        }

        public virtual void CopyTo3(int index, Array array, int arrayIndex, int count)
        {
            _arrayList.CopyTo(index, array, arrayIndex, count);
        }

        public void CopyTo(Array array, int index)
        {
            ((GCollections.ICollection)_arrayList).CopyTo(array, index);
        }

        public static ArrayList FixedSize(ArrayList list)
        {
            return new ArrayList(GCollections.ArrayList.FixedSize(list.WrappedArrayList));

        }
        public static GCollections.IList FixedSize(GCollections.IList list)
        {
            return GCollections.ArrayList.FixedSize(list);
        }

        public virtual GCollections.IEnumerator GetEnumerator()
        { 
            return _arrayList.GetEnumerator(); 
        }

        public virtual GCollections.IEnumerator GetEnumerator2(int index, int count)
        {
            return _arrayList.GetEnumerator(index,count);
        }

        public virtual ArrayList GetRange(int index, int count)
        {
            return new ArrayList(_arrayList.GetRange(index, count));
        }

        public virtual int IndexOf(object value)
        {
            return _arrayList.IndexOf(value);
        }

        public virtual int IndexOf2(object value, int startIndex)
        {
            return _arrayList.IndexOf(value, startIndex);
        }

        public virtual int IndexOf3(object value, int startIndex, int count)
        {
            return _arrayList.IndexOf(value, startIndex, count);
        }

        public virtual void Insert(int index, object value)
        {
            _arrayList.Insert(index, value);
        }

        public virtual void InsertRange(int index, GCollections.ICollection c)
        {
            _arrayList.InsertRange(index, c);
        }

        public virtual int LastIndexOf(object value)
        {
            return _arrayList.LastIndexOf(value);
        }

        public virtual int LastIndexOf2(object value, int startIndex)
        {
            return _arrayList.LastIndexOf(value,startIndex); 
        }

        public virtual int LastIndexOf3(object value, int startIndex, int count)
        {
            return _arrayList.LastIndexOf(value,startIndex,count);
        }

        public static ArrayList ReadOnly(ArrayList list)
        {
            return new ArrayList(GCollections.ArrayList.ReadOnly(list.WrappedArrayList));
        }

        public static GCollections.IList ReadOnly(GCollections.IList list)
        {
            return new ArrayList(GCollections.ArrayList.ReadOnly(list));
        }


        public virtual void Remove(object value)
        {
            _arrayList.Remove(value);
        }

        public virtual void RemoveAt(int index)
        {
            _arrayList.RemoveAt(index);
        }

        public virtual void RemoveRange(int index, int count)
        {
            _arrayList.RemoveRange(index, count);
        }

        public static ArrayList Repeat(object value, int count)
        {
            return new ArrayList(GCollections.ArrayList.Repeat(value, count));
        }

        public virtual void Reverse()
        { 
            _arrayList.Reverse();
        }

        public virtual void Reverse2(int index, int count)
        {
            _arrayList.Reverse(index, count);
        }

        public virtual void SetRange(int index, GCollections.ICollection c)
        {
            _arrayList.SetRange(index, c);
        }

        public virtual void Sort()
        { 
            _arrayList.Sort(); 
        }

        public virtual void Sort2(GCollections.IComparer comparer)
        { 
            _arrayList.Sort(comparer);
        }

        public virtual void Sort3(int index, int count, GCollections.IComparer comparer)
        {
            _arrayList.Sort(index,count,comparer);
        }

        public static ArrayList Synchronized(ArrayList list)
        {
            return new ArrayList(GCollections.ArrayList.Synchronized(list.WrappedArrayList));
        }

        public static GCollections.IList Synchronized(GCollections.IList list)
        {
            return GCollections.ArrayList.Synchronized(list);
        }

        public virtual object[] ToArray()
        { 
            return _arrayList.ToArray(); 
        }

        public virtual Array ToArray2(Type type)
        {
            return _arrayList.ToArray(type);
        }

        public virtual void TrimToSize()
        { 
            _arrayList.TrimToSize();
        }



    }
}
