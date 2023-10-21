using GCollections = global::System.Collections;
using System.Collections;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace DotNetLib.System.Collections
{
    //
    // Summary:
    //     Represents a nongeneric collection of key/value pairs.
    [ComVisible(true)]
    [Guid("7C303967-4125-474D-B3E1-139A183A4156")]
    [Description("Represents a nongeneric collection of key/value pairs.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDictionary : ICollection, GCollections.IEnumerable//, GCollections.IDictionary
    {
        //
        // Summary:
        //     Gets or sets the element with the specified key.
        //
        // Parameters:
        //   key:
        //     The key of the element to get or set.
        //
        // Returns:
        //     The element with the specified key, or null if the key does not exist.
        //
        // Exceptions:
        //   T:System.ArgumentNullException:
        //     key is null.
        //
        //   T:System.NotSupportedException:
        //     The property is set and the System.Collections.IDictionary object is read-only.
        //     -or- The property is set, key does not exist in the collection, and the System.Collections.IDictionary
        //     has a fixed size.
        //[__DynamicallyInvokable]
        object this[object key]
        {
            //[__DynamicallyInvokable]
            get;
            //[__DynamicallyInvokable]
            set;
        }

        //
        // Summary:
        //     Gets an System.Collections.ICollection object containing the keys of the System.Collections.IDictionary
        //     object.
        //
        // Returns:
        //     An System.Collections.ICollection object containing the keys of the System.Collections.IDictionary
        //     object.
        //[__DynamicallyInvokable]
        ICollection Keys
        {
            //[__DynamicallyInvokable]
            get;
        }

        //
        // Summary:
        //     Gets an System.Collections.ICollection object containing the values in the System.Collections.IDictionary
        //     object.
        //
        // Returns:
        //     An System.Collections.ICollection object containing the values in the System.Collections.IDictionary
        //     object.
        //[__DynamicallyInvokable]
        ICollection Values
        {
            //[__DynamicallyInvokable]
            get;
        }

        //
        // Summary:
        //     Gets a value indicating whether the System.Collections.IDictionary object is
        //     read-only.
        //
        // Returns:
        //     true if the System.Collections.IDictionary object is read-only; otherwise, false.
        //[__DynamicallyInvokable]
        bool IsReadOnly
        {
            //[__DynamicallyInvokable]
            get;
        }

        //
        // Summary:
        //     Gets a value indicating whether the System.Collections.IDictionary object has
        //     a fixed size.
        //
        // Returns:
        //     true if the System.Collections.IDictionary object has a fixed size; otherwise,
        //     false.
        //[__DynamicallyInvokable]
        bool IsFixedSize
        {
            //[__DynamicallyInvokable]
            get;
        }

        //
        // Summary:
        //     Determines whether the System.Collections.IDictionary object contains an element
        //     with the specified key.
        //
        // Parameters:
        //   key:
        //     The key to locate in the System.Collections.IDictionary object.
        //
        // Returns:
        //     true if the System.Collections.IDictionary contains an element with the key;
        //     otherwise, false.
        //
        // Exceptions:
        //   T:System.ArgumentNullException:
        //     key is null.
        //[__DynamicallyInvokable]
        bool Contains(object key);

        //
        // Summary:
        //     Adds an element with the provided key and value to the System.Collections.IDictionary
        //     object.
        //
        // Parameters:
        //   key:
        //     The System.Object to use as the key of the element to add.
        //
        //   value:
        //     The System.Object to use as the value of the element to add.
        //
        // Exceptions:
        //   T:System.ArgumentNullException:
        //     key is null.
        //
        //   T:System.ArgumentException:
        //     An element with the same key already exists in the System.Collections.IDictionary
        //     object.
        //
        //   T:System.NotSupportedException:
        //     The System.Collections.IDictionary is read-only. -or- The System.Collections.IDictionary
        //     has a fixed size.
        //[__DynamicallyInvokable]
        void Add(object key, object value);

        //
        // Summary:
        //     Removes all elements from the System.Collections.IDictionary object.
        //
        // Exceptions:
        //   T:System.NotSupportedException:
        //     The System.Collections.IDictionary object is read-only.
        //[__DynamicallyInvokable]
        void Clear();

        //
        // Summary:
        //     Returns an System.Collections.IDictionaryEnumerator object for the System.Collections.IDictionary
        //     object.
        //
        // Returns:
        //     An System.Collections.IDictionaryEnumerator object for the System.Collections.IDictionary
        //     object.
        //[__DynamicallyInvokable]
        new IDictionaryEnumerator GetEnumerator();

        //
        // Summary:
        //     Removes the element with the specified key from the System.Collections.IDictionary
        //     object.
        //
        // Parameters:
        //   key:
        //     The key of the element to remove.
        //
        // Exceptions:
        //   T:System.ArgumentNullException:
        //     key is null.
        //
        //   T:System.NotSupportedException:
        //     The System.Collections.IDictionary object is read-only. -or- The System.Collections.IDictionary
        //     has a fixed size.
        //[__DynamicallyInvokable]
        void Remove(object key);
    }


}
