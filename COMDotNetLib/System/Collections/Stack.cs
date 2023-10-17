// https://learn.microsoft.com/en-us/dotnet/api/system.collections.stack?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections;
using DotNetLib.Extensions;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a simple last-in-first-out (LIFO) non-generic collection of objects.")]
    [Guid("0BAA63BD-14C1-483D-BD52-663C90F648B0")]
    [ProgId("DotNetLib.System.Collections.Stack")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStack))]
    public class Stack : ICollection, IEnumerable, ICloneable, IWrappedObject, IStack
    {
        private GCollections.Stack _stack;

        // Constructors
        public Stack()
        {
            _stack = new GCollections.Stack();
        }

        public Stack(int initialCapacity)
        {
            _stack = new GCollections.Stack(initialCapacity);
        }

        //
        // Summary:
        //     Initializes a new instance of the System.Collections.Stack class that contains
        //     elements copied from the specified collection and has the same initial capacity
        //     as the number of elements copied.
        //
        // Parameters:
        //   col:
        //     The System.Collections.ICollection to copy elements from.
        //
        // Exceptions:
        //   T:System.ArgumentNullException:
        //     col is null.
        public Stack(ICollection col)
            : this(col?.Count ?? 32)
        {
            if (col == null)
            {
                throw new ArgumentNullException("col");
            }

            IEnumerator enumerator = col.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Push(enumerator.Current);
            }
        }

        internal Stack(GCollections.Stack stack)
        {
            _stack = stack;
        }

        // Properties
        public object WrappedObject => _stack;

        internal GCollections.Stack WrappedStack
        {
            get { return _stack; }
        }

        public int Count => _stack.Count;

        public object SyncRoot => _stack.SyncRoot;

        public bool IsSynchronized => _stack.IsSynchronized;


        // Methods
        public virtual void Clear()
        { 
            _stack.Clear(); 
        }

        public virtual bool Contains(object value)
        {
            return _stack.Contains(value);
        }
        public object Clone()
        {
            return new Stack((GCollections.Stack)_stack.Clone());
        }

        public void CopyTo(Array array, int index)
        {
            _stack.CopyTo(array.WrappedArray, index);
        }

        public IEnumerator GetEnumerator()
        {
            return _stack.GetEnumerator();
        }

        new public Type GetType()
        {
            return new Type(((object)this).GetType());
        }

        public virtual object Peek()
        { 
            return _stack.Peek(); 
        }

        public virtual object Pop()
        { 
            return _stack.Pop(); 
        }

        public virtual void Push(object obj)
        { 
            _stack.Push(obj); 
        }

        public static Stack Synchronized(Stack stack)
        {
            return new Stack(GCollections.Stack.Synchronized(stack.WrappedStack));
        }

        public Array ToArray()
        {
            return new Array(_stack.ToArray());
        }

        public void CopyTo(global::System.Array array, int index)
        {
            throw new NotImplementedException();
        }
    }
}
