using System.Runtime.InteropServices;
using GSystem = global::System; // https://stackoverflow.com/questions/5681537/namespace-conflict-in-c-sharp
using System.ComponentModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a strongly typed list of strings that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [Guid("70EF1DD1-80FA-4276-8051-426C0FAFB2FA")]
    [ClassInterface(ClassInterfaceType.None)]
    public class ListString  : IListString
    {
        private GSystem.Collections.Generic.List<string> objListString;
        public ListString()
        {
            this.objListString = new List<string>();
        }

        public ListString(List<String> objListString)
        {
            this.objListString = objListString;
        }

        public ListString(int capacity)
        {
            this.objListString = new List<string>(capacity);
        }

        public ListString(GSystem.Collections.Generic.IEnumerable<string> collection)
        {
            this.objListString = new List<string>(collection);
        }

        public ListString Create()
        {
            return new ListString();
        }

        public ListString Create2(int capacity)
        {
            return new ListString(capacity);
        }

        public ListString CreateFromIEnumerable(GSystem.Collections.IEnumerable collection)
        {
            return new ListString((IEnumerable<string>)collection);
        }

        // Properties
        public int Capacity
        {
            get { return this.objListString.Capacity; }
            set { this.objListString.Capacity = value; }
        }

        public int Count
        {
            get { return this.objListString.Count; }
        }

        public string this[int index]
        {
            get { return this.objListString[index]; }
            set { this.objListString[index] = value; }
        }

        // Methods
        public void Add(string value)
        {
            this.objListString.Add(value);
        }

        //public void AddRange(GSystem.Collections.Generic.IEnumerable<string> collection)
        //{
        //    this.objListString.AddRange((IEnumerable<string>)collection);
        //}

        //public int BinarySearch(T item);
        //public int BinarySearch (T item, System.Collections.Generic.IComparer<T> comparer);
        //public int BinarySearch (int index, int count, T item, System.Collections.Generic.IComparer<T> comparer);

        public int BinarySearch(string item)
        {
            return this.objListString.BinarySearch(item);
        }

        //public int BinarySearch2(string item, GSystem.Collections.Generic.IComparer<string> comparer)
        //{
        //    return this.objListString.BinarySearch(item, comparer);
        //}

        public void Clear()
        {
            this.objListString.Clear();
        }

        public bool Contains(string value)
        {
            return this.objListString.Contains(value);
        }

        //public System.Collections.Generic.List<TOutput> ConvertAll<TOutput>(Converter<T, TOutput> converter);

        public void CopyTo(string[] array)
        {
            this.objListString.CopyTo(array);
        }

        public void CopyTo2(string[] array, int arrayIndex)
        {
            this.objListString.CopyTo(array, arrayIndex);
        }

        public void CopyTo3(int index, string[] array, int arrayIndex, int count)
        {
            this.objListString.CopyTo(index, array, arrayIndex, count);
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
        //    return this.objListString.GetEnumerator();
        //}

        //public GSystem.Collections.Generic.List<string> GetRange(int index, int count)
        //{
        //    return this.objListString.GetRange(index,count);
        //}

        public int IndexOf(string item)
        {
            return this.objListString.IndexOf(item);
        }

        public int IndexOf2(string item, int index)
        {
            return this.objListString.IndexOf(item, index);
        }

        public int IndexOf3(string item, int index, int count)
        {
            return this.objListString.IndexOf(item, index, count);
        }

        public void Insert(int index, string item)
        {
            this.objListString.Insert(index, item);
        }

        //public void InsertRange(int index, System.Collections.Generic.IEnumerable<T> collection);

        public int LastIndexOf(string item)
        {
            return objListString.LastIndexOf(item);
        }

        public int LastIndexOf2(string item, int index)
        {
            return objListString.LastIndexOf(item, index);
        }

        public int LastIndexOf3(string item, int index, int count)
        {
            return objListString.LastIndexOf(item, index, count);
        }

        public bool Remove(string item)
        {
            return this.objListString.Remove(item);
        }

        //public int RemoveAll (Predicate<T> match);

        public void RemoveAt(int index)
        {
            this.objListString.RemoveAt(index);
        }

        public void RemoveRange(int index, int count)
        {
            this.objListString.RemoveRange(index, count);
        }

        public void Reverse()
        {
            this.objListString.Reverse();
        }

        //public void Sort (Comparison<string> comparison)
        //{ 
        //    this.objListString.Sort(comparison);
        //}

        //public void Sort (int index, int count, System.Collections.Generic.IComparer<T> comparer);
        public void Sort()
        {
            this.objListString.Sort();
        }

        public string[] ToArray()
        { 
            return this.objListString.ToArray();
        }

        public void TrimExcess()
        {
            this.objListString.TrimExcess();
        }

        //public bool TrueForAll(Predicate<T> match);

    }
}
