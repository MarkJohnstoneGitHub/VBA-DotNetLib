// https://learn.microsoft.com/en-us/dotnet/api/system.collections.icollection?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("20DAA84D-8FDC-4779-ACFC-B7FBFFFDF9A2")]
    [Description("Defines size, enumerators, and synchronization methods for all nongeneric collections.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICollection : GCollections.IEnumerable ,  GCollections.ICollection
    {
        //
        // Summary:
        //     Gets the number of elements contained in the System.Collections.ICollection.
        //
        // Returns:
        //     The number of elements contained in the System.Collections.ICollection.
        //[__DynamicallyInvokable]
        new int Count
        {
            //[__DynamicallyInvokable]
            get;
        }

        //
        // Summary:
        //     Gets an object that can be used to synchronize access to the System.Collections.ICollection.
        //
        // Returns:
        //     An object that can be used to synchronize access to the System.Collections.ICollection.
        //[__DynamicallyInvokable]
        new object SyncRoot
        {
            //[__DynamicallyInvokable]
            get;
        }

        //
        // Summary:
        //     Gets a value indicating whether access to the System.Collections.ICollection
        //     is synchronized (thread safe).
        //
        // Returns:
        //     true if access to the System.Collections.ICollection is synchronized (thread
        //     safe); otherwise, false.
        //[__DynamicallyInvokable]
        new bool IsSynchronized
        {
            //[__DynamicallyInvokable]
            get;
        }

        //
        // Summary:
        //     Copies the elements of the System.Collections.ICollection to an System.Array,
        //     starting at a particular System.Array index.
        //
        // Parameters:
        //   array:
        //     The one-dimensional System.Array that is the destination of the elements copied
        //     from System.Collections.ICollection. The System.Array must have zero-based indexing.
        //
        //   index:
        //     The zero-based index in array at which copying begins.
        //
        // Exceptions:
        //   T:System.ArgumentNullException:
        //     array is null.
        //
        //   T:System.ArgumentOutOfRangeException:
        //     index is less than zero.
        //
        //   T:System.ArgumentException:
        //     array is multidimensional. -or- The number of elements in the source System.Collections.ICollection
        //     is greater than the available space from index to the end of the destination
        //     array. -or- The type of the source System.Collections.ICollection cannot be cast
        //     automatically to the type of the destination array.
        //[__DynamicallyInvokable]
        void CopyTo(Array array, int index);
    }
}
