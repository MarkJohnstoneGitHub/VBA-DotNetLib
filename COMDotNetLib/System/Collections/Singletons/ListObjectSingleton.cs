// https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a list of <objects> that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [Guid("1A8704BA-27AA-4F71-9AE9-3FC966810688")]
    [ProgId("DotNetLib.System.Collections.ListObjectSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IListObjectSingleton))]
    public class ListObjectSingleton : IListObjectSingleton
    {
        //public ListObjectSingleton() { }

        public ListObject Create()
        {
            return new ListObject();
        }
        public ListObject Create2(int capacity)
        {
            return new ListObject(capacity);
        }

        public ListObject CreateFromIEnumerable(GCollections.IEnumerable collection)
        {
            return new ListObject((IEnumerable<string>)collection);
        }


    }
}
