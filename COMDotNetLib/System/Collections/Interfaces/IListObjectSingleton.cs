// https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("08DE6F0B-5828-4BD6-9A65-47A8142B1C9A")]
    [Description("Represents a list of objects that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IListObjectSingleton
    {
        [Description("Initializes a new instance of the List<Object> class that is empty and has the default initial capacity.")]
        ListObject Create();

        [Description("Initializes a new instance of the List<Object> class that is empty and has the default or specified initial capacity.")]
        ListObject Create2(int capacity);

        [Description("Initializes a new instance of the List<Object> class that contains elements copied from the specified collection and has sufficient capacity to accommodate the number of elements copied.")]
        ListObject CreateFromIEnumerable(GCollections.IEnumerable collection);

    }
}
