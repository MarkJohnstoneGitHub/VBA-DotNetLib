// https://learn.microsoft.com/en-us/dotnet/api/system.collections.queue?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{

    [ComVisible(true)]
    [Description("Represents a first-in, first-out collection of objects.")]
    [Guid("3E9FB6AA-E72A-4B26-80A0-DC93986AE526")]
    [ProgId("DotNetLib.System.Collections.QueueSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IQueueSingleton))]
    public class QueueSingleton : IQueueSingleton
    {
        public QueueSingleton() { }

        //public Queue Create()
        //{
        //    return new Queue();
        //}

        //public Queue Create(int capacity)
        //{ 
        //    return new Queue(capacity); 
        //}

        //public Queue Create(int capacity, float growFactor)
        //{
        //    return new Queue(capacity, growFactor);
        //}

        public Queue Create(int capacity = 32, float growFactor = 2f)
        { 
            return new Queue(capacity, growFactor);
        }

        public Queue Create2(ICollection col)
        {
            return new Queue(col);
        }

        public Queue Synchronized(Queue queue)
        {
            return Queue.Synchronized(queue);
        }
    }
}
