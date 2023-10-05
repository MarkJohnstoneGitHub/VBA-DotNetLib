// https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;
using GCollections = global::System.Collections;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Implements the IList interface using an array whose size is dynamically increased as required.")]
    [Guid("1DFBFA78-E41B-432E-9F19-3DA28D94FD75")]
    [ProgId("DotNetLib.System.Collections.ArrayListSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IArrayListSingleton))]

    public class ArrayListSingleton : IArrayListSingleton
    {
        public ArrayListSingleton() {}

        // Factory Methods
        public ArrayList Create(int capacity = 0)
        {
            return new ArrayList(capacity);
        }

        public ArrayList Create2(GCollections.ICollection c)
        {
            return new ArrayList(c);
        }

        public ArrayList Adapter(GCollections.IList list)
        { 
            return ArrayList.Adapter(list); 
        }

        public ArrayList FixedSize(ArrayList list)
        {
            return ArrayList.FixedSize(list);
        }

        public GCollections.IList FixedSize2(GCollections.IList list)
        {
            return ArrayList.FixedSize(list);
        }

        public ArrayList ReadOnly(ArrayList list)
        {
            return ArrayList.ReadOnly(list);
        }

        public GCollections.IList ReadOnly2(GCollections.IList list)
        {
            return ArrayList.ReadOnly(list);
        }

        public ArrayList Repeat(object value, int count)
        {
            return ArrayList.Repeat(value, count);
        }

        public ArrayList Synchronized(ArrayList list)
        {
            return ArrayList.Synchronized(list);
        }

        public GCollections.IList Synchronized2(GCollections.IList list)
        {
            return ArrayList.Synchronized(list);
        }

    }
}
