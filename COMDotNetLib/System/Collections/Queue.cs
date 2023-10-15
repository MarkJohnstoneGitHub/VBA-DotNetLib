// https://learn.microsoft.com/en-us/dotnet/api/system.collections.queue?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System;
using System.Collections;
using DotNetLib.Extensions;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a first-in, first-out collection of objects.")]
    [Guid("DB73D389-63F4-4800-9F0F-8A0DB8958AAC")]
    [ProgId("DotNetLib.System.Collections.Queue")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IQueue))]
    public class Queue : ICollection, GCollections.IEnumerable, ICloneable, IQueue, IWrappedObject
    {
        private GCollections.Queue _queue;

        // Constructors

        internal Queue(GCollections.Queue queue)
        {
            _queue = queue;
        }

        public Queue()
        {
            _queue = new GCollections.Queue();
        }

        public Queue(GCollections.ICollection col)
        {
            _queue = new GCollections.Queue(col);
        }

        //
        // Summary:
        //     Initializes a new instance of the System.Collections.Queue class that contains
        //     elements copied from the specified collection, has the same initial capacity
        //     as the number of elements copied, and uses the default growth factor.
        //
        // Parameters:
        //   col:
        //     The System.Collections.ICollection to copy elements from.
        //
        // Exceptions:
        //   T:System.ArgumentNullException:
        //     col is null.
        public Queue(ICollection col)
            : this(col?.Count ?? 32)
        {
            if (col == null)
            {
                throw new ArgumentNullException("col");
            }

            IEnumerator enumerator = col.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Enqueue(enumerator.Current);
            }
        }


        public Queue(int capacity)
        {
            _queue = new GCollections.Queue(capacity);
        }

        public Queue(int capacity, float growFactor)
        {
            _queue = new GCollections.Queue(capacity, growFactor);
        }

        // Properties

        public object WrappedObject => _queue;

        internal GCollections.Queue WrappedQueue
        {
            get { return _queue; }
        }

        public int Count => _queue.Count;

        public object SyncRoot => _queue.SyncRoot;

        public bool IsSynchronized => _queue.IsSynchronized;

        // Methods

        public void Clear()
        {
            _queue.Clear();
        }

        public object Clone()
        {
            return new Queue((GCollections.Queue)_queue.Clone());
        }

        public bool Contains(object obj)
        {
            return _queue.Contains(obj);
        }

        public void CopyTo(Array array, int index)
        {
            _queue.CopyTo(array.WrappedArray, index);
        }

        public object Dequeue()
        {
            return _queue.Dequeue();
        }

        public void Enqueue(object obj)
        {
            _queue.Enqueue(obj);
        }

        public IEnumerator GetEnumerator()
        {
            return _queue.GetEnumerator();
        }

        public object Peek()
        {
            return _queue.Peek();
        }

        public static Queue Synchronized(Queue queue)
        {
            return new Queue(GCollections.Queue.Synchronized(queue.WrappedQueue));
        }

        //public object[] ToArray()
        //{
        //    return _queue.ToArray();
        //}

        public Array ToArray()
        {
            return new Array(_queue.ToArray());
        }


        public void TrimToSize()
        {
            _queue.TrimToSize();
        }

        new public virtual bool Equals(object obj)
        { 
            return _queue.Equals(obj.Unwrap()); 
        }


    }
}
