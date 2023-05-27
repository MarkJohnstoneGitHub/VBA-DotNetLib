// https://stackoverflow.com/questions/1296362/how-to-expose-a-dictionary-to-com-interop?rq=3
// https://stackoverflow.com/questions/17519078/initializing-a-generic-variable-from-a-c-sharp-type-variable

using GSystem = global::System;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{

    [ComVisible(true)]
    [Description("Represents a strongly typed list of objects that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [Guid("C88C9749-4D9C-46D0-A463-5DA93F0E1A75")]
    //[ProgId("DotNetLib.System.Collections.Generic.List")]
    [ClassInterface(ClassInterfaceType.None)]


    public class List : IList
    {
        private GSystem.Collections.Generic.List<object> objList;

        // Constructors
        public List()
        {
            GSystem.Collections.Generic.List<object> objList = new GSystem.Collections.Generic.List<object>();
        }

        public List(List<object> objList)
        {
            this.objList = (List<object>)objList;
        }

        public List(GSystem.Collections.Generic.List<object> objList, int capacity)
        {
            this.objList = objList;
            this.objList.Capacity = capacity;
        }

        public List(GSystem.Collections.Generic.IEnumerable<object> collection)
        {
            this.objList = new List<object>(collection);
        }

        // https://stackoverflow.com/questions/17519078/initializing-a-generic-variable-from-a-c-sharp-type-variable
        // https://learn.microsoft.com/en-us/dotnet/api/system.type.makegenerictype?redirectedfrom=MSDN&view=net-7.0#System_Type_MakeGenericType_System_Type___
        public List Create(object listType)
        {
            Type type = typeof(GSystem.Collections.Generic.List<>).MakeGenericType(listType.GetType());
            object objType = Activator.CreateInstance(type);
            return new List((List<object>)objType);
        }

        public List Create2(object listType, int capacity)
        {
            Type type = typeof(GSystem.Collections.Generic.List<>).MakeGenericType(listType.GetType());
            object objType = Activator.CreateInstance(type);
            return new List((List<object>)objType,capacity);
        }

        public List CreateFromIEnumerable(GSystem.Collections.IEnumerable collection)
        {
            return new List((IEnumerable<object>)collection);
        }

        // Properties
        public int Capacity
        { 
            get { return this.objList.Capacity; }
            set { this.objList.Capacity = value; }
        }
        public int Count
        {
            get { return this.objList.Count; }
        }

        public object this[int index]
        {
            get { return this.objList[index]; }
            set { this.objList[index] = value; }
        }

        // Methods
        public void Add(object value)
        { 
            this.objList.Add(value); 
        }

        public void AddRange(GSystem.Collections.IEnumerable collection)
        {
            this.objList.AddRange((IEnumerable<object>)collection);
        }

        //public int BinarySearch(T item);
        //public int BinarySearch (T item, System.Collections.Generic.IComparer<T> comparer);
        //public int BinarySearch (int index, int count, T item, System.Collections.Generic.IComparer<T> comparer);

        public void Clear()
        {
            this.objList.Clear(); 
        }

        public bool Contains(object value)
        {
            return this.objList.Contains(value);
        }

        //public System.Collections.Generic.List<TOutput> ConvertAll<TOutput>(Converter<T, TOutput> converter);

        public void CopyTo(object[] array)
        {
            this.objList.CopyTo(array);
        }

        public void CopyTo2(object[] array, int arrayIndex)
        {
            this.objList.CopyTo(array, arrayIndex);
        }

        public void CopyTo3(int index, object[] array, int arrayIndex, int count)
        {
            this.objList.CopyTo(index, array, arrayIndex, count);
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

        //public System.Collections.Generic.List<T>.Enumerator GetEnumerator();

        //public System.Collections.Generic.List<T> GetRange(int index, int count);

        //public int IndexOf (T item);
        //public int IndexOf(T item, int index);
        //public int IndexOf(T item, int index, int count);
        public int IndexOf(object item)
        { 
            return this.objList.IndexOf(item);
        }

        public int IndexOf2(object item, int index)
        {
            return this.objList.IndexOf(item,index);
        }

        public int IndexOf3(object item, int index, int count)
        {
            return this.objList.IndexOf(item, index,count);
        }


        public void Insert(int index, object item)
        {
            this.objList.Insert(index, item); 
        }

        //public void InsertRange(int index, System.Collections.Generic.IEnumerable<T> collection);

        public int LastIndexOf(object item)
        {
            return objList.LastIndexOf(item);
        }

        public int LastIndexOf2(object item, int index)
        {
            return objList.LastIndexOf(item,index);
        }

        public int LastIndexOf3(object item, int index, int count)
        {
            return objList.LastIndexOf(item, index,count);
        }

        public bool Remove(object item)
        { 
            return this.objList.Remove(item);
        }

        //public int RemoveAll (Predicate<T> match);

        public void RemoveAt(int index)
        { 
            this.objList.RemoveAt(index);
        }

        public void RemoveRange(int index, int count)
        {
            this.objList.RemoveRange(index, count);
        }

        public void Reverse()
        { 
            this.objList.Reverse(); 
        }

        //public void Sort (Comparison<T> comparison);
        //public void Sort (int index, int count, System.Collections.Generic.IComparer<T> comparer);
        public void Sort()
        { 
            this.objList.Sort(); 
        }


        public object[] ToArray()
        { 
            return this.objList.ToArray(); 
        }

        public void TrimExcess()
        { 
            this.objList.TrimExcess();
        }

        //public bool TrueForAll(Predicate<T> match);
    }
}
