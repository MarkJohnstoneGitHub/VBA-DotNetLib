// https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a strongly typed list of <strings> that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [Guid("B55933C0-FBA6-4692-AE40-F26D6C2E1ED5")]
    [ProgId("DotNetLib.System.Collections.ListStringSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IListStringSingleton))]
    public class ListStringSingleton  : IListStringSingleton
    {
        public ListString Create()
        {
            return new ListString();
        }
        public ListString Create2(int capacity)
        {
            return new ListString(capacity);
        }

        public ListString CreateFromIEnumerable(GCollections.IEnumerable collection)
        {
            return new ListString((IEnumerable<string>)collection);
        }
    }
}
