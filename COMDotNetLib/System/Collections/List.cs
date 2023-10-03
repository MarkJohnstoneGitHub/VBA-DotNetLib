// https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1?view=netframework-4.8.1

// Notes:
// https://stackoverflow.com/questions/9860387/how-do-i-create-a-dynamic-type-listt
// https://stackoverflow.com/questions/906499/getting-type-t-from-ienumerablet
//   https://stackoverflow.com/a/17713382/10759363
//   https://stackoverflow.com/a/57679532/10759363

// https://stackoverflow.com/questions/1296362/how-to-expose-a-dictionary-to-com-interop?rq=3
// https://stackoverflow.com/questions/17519078/initializing-a-generic-variable-from-a-c-sharp-type-variable

// https://stackoverflow.com/a/254496/10759363

using GGeneric = global::System.Collections.Generic;
using GCollections = global::System.Collections;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a strongly typed list of objects that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [Guid("C88C9749-4D9C-46D0-A463-5DA93F0E1A75")]
    [ProgId("DotNetLib.System.Collections.List")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IList))]
    public class List : GCollections.IList, ICollection, IEnumerable , IList 
    {
        private GCollections.IList _list;

        public List()
        {
            _list = new List<object>();
        }

        public List(object type) 
        {
            _list = CreateFromType((dynamic)type);
        }

        public List(object type, int capacity)
        {
            _list = CreateFromType(capacity, (dynamic)type);
        }

        //public List(IEnumerable collection)
        //{
        //    _list = new List<object>((IEnumerable<object>)collection);
        //}

        //public List(object type, IEnumerable collection)
        //{
        //    GCollections.IList list = CreateFromType((dynamic)type);
        //    // require to check types are the same?
        //    // Do you require the type?
        //    _list = new List<object>((IEnumerable<object>)collection);
        //}

        internal static List<T> CreateFromType<T>(T obj = default)
        {
            return new List<T>();
        }

        internal static List<T> CreateFromType<T>(int capacity, T obj = default)
        {
            return new List<T>(capacity);
        }

        // Properties

        internal List<object> GenericList => (List<object>)_list;

        public bool IsReadOnly => _list.IsReadOnly;

        public bool IsFixedSize => _list.IsFixedSize;

        public int Count => _list.Count;

        public object SyncRoot => _list.SyncRoot;

        public bool IsSynchronized => _list.IsSynchronized;

        public int Capacity 
        {
            get => GenericList.Capacity;
            set => GenericList.Capacity = value;
        }

        public object this[int index]
        {
            get => _list[index];
            set => _list[index] = value;
        }

        // Methods
        public void Add(object item)
        {
            _list.Add(item);
        }

        public void AddRange(IEnumerable collection)
        {
            GenericList.Add(collection);
        }

        public int BinarySearch(object item)
        { 
            return GenericList.BinarySearch(item);
        }

        public int BinarySearch(object item, IComparer comparer)
        {
            return GenericList.BinarySearch(item, (IComparer<object>)comparer);

        }

        public void Clear()
        {
            _list.Clear();
        }

        public bool Contains(object value)
        {
            return _list.Contains(value);
        }

        //public System.Collections.Generic.List<TOutput> ConvertAll<TOutput>(Converter<T, TOutput> converter);

        public void CopyTo(Array array, int index)
        {
            _list.CopyTo(array, index);
        }

        //https://stackoverflow.com/questions/68481139/how-to-convert-my-predicate-to-a-generic-predicate-in-c
        //https://stackoverflow.com/questions/9842222/dynamic-cast-to-generic-type
        //https://stackoverflow.com/questions/42902164/how-to-map-expressionfunctentity-bool-to-expressionfunctdbentity-bool/42904029#42904029
        //public bool Exists(Predicate<T> match);

        //public T Find(Predicate<T> match);

        //public System.Collections.Generic.List<T> FindAll(Predicate<T> match);

        public IEnumerator GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        public int IndexOf(object value)
        {
            return _list.IndexOf(value);
        }

        public int IndexOf2(object item, int index)
        {
            return GenericList.IndexOf(item, index);
        }

        public int IndexOf3(object item, int index, int count)
        {
            return GenericList.IndexOf(item, index,count);
        }

        public void Insert(int index, object value)
        {
            _list.Insert(index, value);
        }

        public int LastIndexOf(object item)
        {
            return GenericList.LastIndexOf(item);
            
        }

        public int LastIndexOf2(object item, int index)
        {
            return GenericList.LastIndexOf(item, index);
        }

        public int LastIndexOf3(object item, int index, int count)
        {
            return  GenericList.LastIndexOf(item, index,count);
        }

        public bool Remove(object value)
        {
            return GenericList.Remove(value);
        }

        //public int RemoveAll (Predicate<T> match);

        public void RemoveAt(int index)
        {
            _list.RemoveAt(index);
        }

        int GCollections.IList.Add(object value)
        {
            return _list.Add(value);
            
        }

        public void CopyTo(object[] array)
        {
            throw new NotImplementedException();
        }

        public void CopyTo2(object[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public void CopyTo3(int index, object[] array, int arrayIndex, int count)
        {
            throw new NotImplementedException();
        }

        public void RemoveRange(int index, int count)
        {
            GenericList.RemoveRange(index, count);
        }

        public void Reverse()
        {
            GenericList.Reverse();
        }

        public void Sort()
        {
            GenericList.Sort();
        }

        public object[] ToArray()
        {
            throw new NotImplementedException();
        }

        public void TrimExcess()
        {
            GenericList.TrimExcess();
        }

        void GCollections.IList.Remove(object value)
        {
            GenericList.Remove(value);
        }

        //public bool TrueForAll (Predicate<T> match);   

        public static GCollections.IList CreateFromTypeV2<T>(T obj = default)
        {
            Type type = obj.GetType();
            Type listType = typeof(List<>).MakeGenericType(new[] { type });
            GCollections.IList list = (GCollections.IList)Activator.CreateInstance(listType);
            return list;
        }
    }
}
