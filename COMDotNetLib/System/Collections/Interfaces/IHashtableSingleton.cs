// https://learn.microsoft.com/en-us/dotnet/api/system.collections.hashtable?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("10E94F03-C7E1-4CB4-8542-D53D4D3CBCDB")]
    [Description("Represents a collection of key/value pairs that are organized based on the hash code of the key.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IHashtableSingleton
    {
        // [Description("")]

        // Factory methods
        [Description("Initializes a new, empty instance of the Hashtable class using the specified initial capacity and load factor, and the default hash code provider and comparer.")]
        Hashtable Create(int capacity = 0, float loadFactor = 1f, GCollections.IEqualityComparer equalityComparer = null);

        //Hashtable Create2(int capacity = 0, float loadFactor = 1f);

        [Description("Initializes a new instance of the Hashtable class by copying the elements from the specified dictionary to the new Hashtable object. The new Hashtable object has an initial capacity equal to the number of elements copied, and uses the specified load factor and IEqualityComparer object.")]
        Hashtable Create2(GCollections.IDictionary d, float loadFactor = 1f, GCollections.IEqualityComparer equalityComparer = null);

    }
}
