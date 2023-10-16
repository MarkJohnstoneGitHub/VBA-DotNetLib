// https://learn.microsoft.com/en-us/dotnet/api/system.collections.stack?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a simple last-in-first-out (LIFO) non-generic collection of objects.")]
    [Guid("48D47A84-4A68-4F8D-BD41-61FD23290A54")]
    [ProgId("DotNetLib.System.Collections.StackSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStackSingleton))]
    public class StackSingleton : IStackSingleton
    {
        public Stack Create(int initialCapacity = 10)
        {
            return new Stack(initialCapacity);
        }

        public Stack Create2(ICollection col)
        {
            return new Stack(col);
        }

        public Stack Synchronized(Stack col)
        {
            return Stack.Synchronized(col);
        }
    }
}
