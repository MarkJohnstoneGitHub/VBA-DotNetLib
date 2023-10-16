// https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;
using GCollections = global::System.Collections;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("A42CCD15-2E3F-4CCF-87C8-E298B3D31EAF")]
    [Description("Represents a collection of key/value pairs that are sorted by the keys and are accessible by key and by index.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISortedListSingleton
    {
        [Description("Initializes a new instance of the SortedList class that is empty, has the specified initial capacity, and is sorted according to the IComparable interface implemented by each key added to the SortedList object.")]
        SortedList Create(int initialCapacity = 0);

        [Description("Initializes a new instance of the SortedList class that is empty, has the specified initial capacity, and is sorted according to the specified IComparer interface.")]
        SortedList Create2(GCollections.IComparer comparer, int capacity = 0);

        [Description("Initializes a new instance of the SortedList class that contains elements copied from the specified dictionary, has the same initial capacity as the number of elements copied, and is sorted according to the specified IComparer interface.")]
        SortedList Create3(IDictionary d, GCollections.IComparer comparer = null);

    }
}
