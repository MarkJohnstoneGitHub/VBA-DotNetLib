// https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a collection of key/value pairs that are sorted by the keys and are accessible by key and by index.")]
    [Guid("32A6D98E-68D3-4D0D-959E-0737FCE4ADD0")]
    [ProgId("DotNetLib.System.Collections.SortedListSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ISortedListSingleton))]
    public class SortedListSingleton : ISortedListSingleton
    {
        public SortedList Create(int initialCapacity = 0)
        {
            return new SortedList(initialCapacity);
        }

        public SortedList Create2(GCollections.IComparer comparer, int capacity = 0)
        { 
            return new SortedList(comparer, capacity); 
        }

        public SortedList Create3(IDictionary d, GCollections.IComparer comparer = null)
        {
            return new SortedList((GCollections.IDictionary)d, comparer);
        }

    }
}
