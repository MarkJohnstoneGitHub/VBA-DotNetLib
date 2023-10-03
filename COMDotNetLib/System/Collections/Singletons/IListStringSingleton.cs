﻿// https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System.ComponentModel;

namespace DotNetLib.System.Collections
{
    public interface IListStringSingleton
    {
        // Constructors
        //[Description("Initializes a new instance of the List<string> class that is empty and has the default initial capacity.")]
        //ListString Create();

        [Description("Initializes a new instance of the List<string> class that is empty and has the default initial capacity.")]
        ListString Create();

        [Description("Initializes a new instance of the List<string> class that is empty and has the default or specified initial capacity.")]
        ListString Create2(int capacity);

        [Description("Initializes a new instance of the List<string> class that contains elements copied from the specified collection and has sufficient capacity to accommodate the number of elements copied.")]
        ListString CreateFromIEnumerable(GCollections.IEnumerable collection);

    }
}