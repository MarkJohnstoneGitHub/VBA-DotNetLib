// https://learn.microsoft.com/en-us/dotnet/api/system.collections.hashtable?view=netframework-4.8.1

using System;
using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using GCollections = global::System.Collections;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a collection of key/value pairs that are organized based on the hash code of the key.")]
    [Guid("DBC5EB31-393D-4867-AF8C-83C0AC402212")]
    [ProgId("DotNetLib.System.Collections.Hashtable")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IHashtable))]
    public class Hashtable : GCollections.IDictionary, GCollections.ICollection, IEnumerable, ISerializable, IDeserializationCallback, ICloneable, IHashtable
    {
        private GCollections.Hashtable _hashtable;

        // Constructors
        public Hashtable()
        {
            _hashtable = new GCollections.Hashtable();
        }

        public Hashtable(int capacity)
        {
            _hashtable = new GCollections.Hashtable(capacity);
        }

        public Hashtable(int capacity, float loadFactor)
        {
            _hashtable = new GCollections.Hashtable(capacity, loadFactor);
        }

        public Hashtable(GCollections.IEqualityComparer equalityComparer)
        {
            _hashtable = new GCollections.Hashtable(equalityComparer);
        }

        public Hashtable(int capacity, GCollections.IEqualityComparer equalityComparer)
        {
            _hashtable = new GCollections.Hashtable(capacity, equalityComparer);
        }

        public Hashtable(int capacity, float loadFactor, GCollections.IEqualityComparer equalityComparer)
        {
            _hashtable = new GCollections.Hashtable(capacity, loadFactor, equalityComparer);
        }

        public Hashtable(GCollections.IDictionary d)
        {
            _hashtable = new GCollections.Hashtable(d);
        }

        public Hashtable(GCollections.IDictionary d, float loadFactor)
        {
            _hashtable = new GCollections.Hashtable(d, loadFactor);
        }

        public Hashtable(GCollections.IDictionary d, GCollections.IEqualityComparer equalityComparer)
        {
            _hashtable = new GCollections.Hashtable(d, equalityComparer);
        }

        public Hashtable(GCollections.IDictionary d, float loadFactor, GCollections.IEqualityComparer equalityComparer)
        {
            _hashtable = new GCollections.Hashtable(d, loadFactor, equalityComparer);
        }

        // Properties

        public int Count => _hashtable.Count;

        public bool IsFixedSize => _hashtable.IsFixedSize;

        public bool IsReadOnly => _hashtable.IsReadOnly;

        public bool IsSynchronized => _hashtable.IsSynchronized;

        public object this[object key]
        {
            get => _hashtable[key];
            set => _hashtable[key] = value;
        }

        public GCollections.ICollection Keys => _hashtable.Keys;

        public object SyncRoot => _hashtable.SyncRoot;

        public GCollections.ICollection Values => _hashtable.Values;


        // Methods


        public void Add(object key, object value)
        {
            _hashtable.Add(key, value);
        }

        public void Clear()
        {
            _hashtable.Clear();
        }

        public bool Contains(object key)
        {
            return _hashtable.Contains(key);
        }

        //public IDictionaryEnumerator GetEnumerator()
        //{
        //    return _hashtable.GetEnumerator();
        //}

        public void CopyTo(global::System.Array array, int index)
        {
            _hashtable.CopyTo(array, index);
        }

        public object Clone()
        {
            return new Hashtable((GCollections.IDictionary)_hashtable.Clone());
        }

        public void CopyTo(Array array, int index)
        {
            _hashtable.CopyTo(array.WrappedArray, index);
        }

        public bool ContainsKey(object key)
        {
            return _hashtable.ContainsKey(key);
        }

        public bool ContainsValue(object value)
        {
            return _hashtable.ContainsValue(value);
        }

        public IDictionaryEnumerator GetEnumerator()
        {
            return _hashtable.GetEnumerator();
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            ((ISerializable)_hashtable).GetObjectData(info, context);
        }

        new public Type GetType()
        {
            return new Type(((object)this).GetType());
        }

        public void OnDeserialization(object sender)
        {
            _hashtable.OnDeserialization(sender);
        }

        public void Remove(object key)
        {
            _hashtable.Remove(key);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable)_hashtable).GetEnumerator();
        }
    }
}
