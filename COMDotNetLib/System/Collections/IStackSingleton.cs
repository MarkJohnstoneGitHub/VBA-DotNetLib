// https://learn.microsoft.com/en-us/dotnet/api/system.collections.stack?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("689FD57E-D585-47D7-B732-0198863AB5DE")]
    [Description("Represents a simple last-in-first-out (LIFO) non-generic collection of objects.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStackSingleton
    {
        [Description("Initializes a new instance of the Stack class that is empty and has the specified initial capacity or the default initial capacity, whichever is greater.")]
        Stack Create(int initialCapacity = 10);

        [Description("Initializes a new instance of the Stack class that contains elements copied from the specified collection and has the same initial capacity as the number of elements copied.")]
        Stack Create2(ICollection col);

        [Description("Returns a synchronized (thread safe) wrapper for the Stack.")]
        Stack Synchronized(Stack col);
    }
}
