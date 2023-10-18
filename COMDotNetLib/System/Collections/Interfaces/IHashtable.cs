// https://learn.microsoft.com/en-us/dotnet/api/system.collections.hashtable?view=netframework-4.8.1

using GSystem = global::System;
using GCollections = global::System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("2CCC25CC-2970-4AF3-AFB5-4E0D4903C235")]
    [Description("Represents a collection of key/value pairs that are organized based on the hash code of the key.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IHashtable
    {
        int Count 
        {
            [Description("Gets the number of key/value pairs contained in the Hashtable.")]
            get;
        }

        bool IsFixedSize 
        {
            [Description("Gets a value indicating whether the Hashtable has a fixed size.")]
            get;
        }

        bool IsReadOnly
        {
            [Description("Gets a value indicating whether the Hashtable is read-only.")]
            get;
        }

        bool IsSynchronized
        {
            [Description("Gets a value indicating whether access to the Hashtable is synchronized (thread safe).")]
            get;
        }

        object this[object key] 
        {
            [Description("Gets or sets the value associated with the specified key.")]
            get;
            [Description("Gets or sets the value associated with the specified key.")]
            set;
        }

        GCollections.ICollection Keys 
        {
            [Description("Gets an ICollection containing the keys in the Hashtable.")]
            get;
        }

        object SyncRoot
        {
            [Description("Gets an object that can be used to synchronize access to the Hashtable.")]
            get;
        }

        GCollections.ICollection Values 
        {
            [Description("Gets an ICollection containing the values in the Hashtable.")]
            get;
        }

        [Description("Adds an element with the specified key and value into the Hashtable.")]
        void Add(object key, object value);

        [Description("Removes all elements from the Hashtable.")]
        void Clear();

        [Description("Creates a shallow copy of the Hashtable.")]
        object Clone();

        [Description("Determines whether the Hashtable contains a specific key.")]
        bool Contains(object key);

        [Description("Determines whether the Hashtable contains a specific key.")]
        bool ContainsKey(object key);

        [Description("Determines whether the Hashtable contains a specific value.")]
        bool ContainsValue(object value);

        [Description("Copies the Hashtable elements to a one-dimensional Array instance at the specified index.")]
        void CopyTo(Array array, int index);

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("Returns an IDictionaryEnumerator that iterates through the Hashtable.")]
        GCollections.IDictionaryEnumerator GetEnumerator();

        //[Description("")]
        //int GetHash(object key);

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        int GetHashCode();

        [Description("Implements the ISerializable interface and returns the data needed to serialize the Hashtable.")]
        void GetObjectData(GSystem.Runtime.Serialization.SerializationInfo info, GSystem.Runtime.Serialization.StreamingContext context);

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Implements the ISerializable interface and raises the deserialization event when the deserialization is complete.")]
        void OnDeserialization(object sender);

        [Description("Removes the element with the specified key from the Hashtable.")]
        void Remove(object key);

        [Description("Returns a string that represents the current object.\r\n\r\n(Inherited from Object)")]
        string ToString();
    }
}
