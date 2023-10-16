using GCollections = global::System.Collections;
using System;

namespace DotNetLib.System.Collections
{
    public class GenericList : IList
    {
        private DynamicList<Type> _genericList;

        //Default is a list of objects
        public GenericList()
        {   
            object obj = new object();
            _genericList = new DynamicList<Type>(obj.GetType());
        }

        public GenericList(object obj)
        {
            _genericList = new DynamicList<Type>(obj.GetType());
        }

        public object this[int index] 
        {
            get => _genericList.List[index];
            set => _genericList.List[index] = (Type)value;
        }

        public int Capacity 
        { 
            get => _genericList.List.Capacity;
            set => _genericList.List.Capacity = value;
        }

        public int Count => _genericList.List.Count;

        public bool IsReadOnly => throw new NotImplementedException();

        public bool IsFixedSize => throw new NotImplementedException();

        public void Add(object value)
        {
            _genericList.List.Add((Type)value);
        }

        public void AddRange(GCollections.IEnumerable collection)
        {
            _genericList.List.AddRange((GCollections.Generic.IEnumerable<Type>)collection);
        }

        public void Clear()
        {
            _genericList.List.Clear();
        }

        public bool Contains(object value)
        {
            return _genericList.List.Contains((Type)value);
        }

        public GCollections.IEnumerator GetEnumerator()
        {
            return _genericList.List.GetEnumerator();
        }

        public int IndexOf(object item)
        {
            return _genericList.List.IndexOf((Type)item);
        }

        public int IndexOf2(object item, int index)
        {
            return _genericList.List.IndexOf((Type)item, index);
        }

        public int IndexOf3(object item, int index, int count)
        {
            return _genericList.List.IndexOf((Type)item, index,count);
        }

        public void Insert(int index, object item)
        {
            _genericList.List.Insert(index, (Type)item);
        }

        public int LastIndexOf(object item)
        {
            return _genericList.List.LastIndexOf((Type)item);
        }

        public int LastIndexOf2(object item, int index)
        {
            return _genericList.List.LastIndexOf((Type)item,index);
        }

        public int LastIndexOf3(object item, int index, int count)
        {
            return _genericList.List.LastIndexOf((Type)item, index,count);
        }

        public bool Remove(object value)
        {
            return _genericList.List.Remove((Type)value);
        }

        public void RemoveAt(int index)
        {
            _genericList.List.RemoveAt(index);
        }

        public void RemoveRange(int index, int count)
        {
            _genericList.List.RemoveRange(index, count);
        }

        public void Reverse()
        {
            throw new NotImplementedException();
        }

        public void Sort()
        {
            throw new NotImplementedException();
        }

        public object[] ToArray()
        {
            throw new NotImplementedException();
        }

        public void TrimExcess()
        {
            throw new NotImplementedException();
        }
    }
}
