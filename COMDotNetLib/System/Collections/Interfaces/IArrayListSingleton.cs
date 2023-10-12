// https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("18ED5BAB-C47C-4E95-B8BC-56B193B265B4")]
    [Description("Implements the IList interface using an array whose size is dynamically increased as required.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IArrayListSingleton
    {
        [Description("Initializes a new instance of the ArrayList class that is empty and has the default or specified initial capacity.")]
        ArrayList Create(int capacity = 0);

        [Description("Initializes a new instance of the ArrayList class that contains elements copied from the specified collection and that has the same initial capacity as the number of elements copied.")]
        ArrayList Create2(GCollections.ICollection c);

        [Description("Creates an ArrayList wrapper for a specific IList.")]
        ArrayList Adapter(GCollections.IList list);

        [Description("Returns an ArrayList wrapper with a fixed size.")]
        ArrayList FixedSize(ArrayList list);

        [Description("Returns an IList wrapper with a fixed size.")]
        GCollections.IList FixedSize2(GCollections.IList list);

        [Description("Returns a read-only ArrayList wrapper.")]
        ArrayList ReadOnly(ArrayList list);

        [Description("Returns a read-only IList wrapper.")]
        GCollections.IList ReadOnly2(GCollections.IList list);

        [Description("Returns an ArrayList whose elements are copies of the specified value.")]
        ArrayList Repeat(object value, int count);

        [Description("Returns an ArrayList wrapper that is synchronized (thread safe).")]
        ArrayList Synchronized(ArrayList list);

        [Description("Returns an IList wrapper that is synchronized (thread safe).")]
        GCollections.IList Synchronized2(GCollections.IList list);
    }
}
