// https://learn.microsoft.com/en-us/dotnet/api/system.collections.queue?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("15B2F1A3-DB0C-453A-83EE-5066A2B54E62")]
    [Description("Represents a first-in, first-out collection of objects.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IQueueSingleton
    {
        [Description("Initializes a new instance of the Queue class that is empty, has the default or specified initial capacity, and uses the default or specified growth factor.")]
        Queue Create(int capacity = 32, float growFactor = 2f);

        [Description("Initializes a new instance of the Queue class that contains elements copied from the specified collection, has the same initial capacity as the number of elements copied, and uses the default growth factor.")]
        Queue Create2(ICollection col);

        [Description("Returns a new Queue that wraps the original queue, and is thread safe.")]
        Queue Synchronized(Queue queue);



    }
}
