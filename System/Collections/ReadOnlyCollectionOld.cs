// https://learn.microsoft.com/en-us/dotnet/api/system.collections.objectmodel.readonlycollection-1?view=netframework-4.8.1
// https://source.dot.net/#Microsoft.Build/ReadOnlyCollection.cs

using GSystem = global::System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Provides the base class for a generic read-only collection.")]
    [Guid("A762D114-7FBE-444B-96E0-53838D20C49D")]
    [ProgId("DotNetLib.System.Collections.ObjectModel.ReadOnlyCollection")]
    [ClassInterface(ClassInterfaceType.None)]

    public class ReadOnlyCollectionOld : IReadOnlyCollectionOld
    {

        private GSystem.Collections.ObjectModel.ReadOnlyCollection<Object> _readOnlyCollection;

        // Constructors
        public ReadOnlyCollectionOld(GSystem.Collections.Generic.IList<Object> list)
        {
            this._readOnlyCollection = (GSystem.Collections.ObjectModel.ReadOnlyCollection<Object>)list;
        }

        public ReadOnlyCollectionOld Create(IList list)
        {
            return new ReadOnlyCollectionOld((GSystem.Collections.Generic.List<Object>)list);
        }

        public int Count => this._readOnlyCollection.Count;

        //public object Item(int index) => this._readOnlyCollection[index];

        public object this[int index] => this._readOnlyCollection[index];


        //GSystem.Collections.Generic.IList<object> Items => _readOnlyCollection.Items; //{ get; }

        //public object[] Items() => this._readOnlyCollection.CopyTo(this._readOnlyCollection);
        //{ get; }

        public bool Contains(Object Index)
        {
            return this._readOnlyCollection.Contains((Object)Index);
        }

        /// <summary>
        /// Copies the entire ReadOnlyCollection<T> to a compatible one-dimensional Array, starting at the specified index of the target array.
        /// </summary>
        /// <param name="array">The one-dimensional Array that is the destination of the elements copied from ReadOnlyCollection<T>. The Array must have zero-based indexing.</param>
        /// <param name="index">The zero-based index in array at which copying begins.</param>
        public void CopyTo(Object[] array, int index)
        {
            this._readOnlyCollection.CopyTo(array, index);
        }

        public GSystem.Collections.IEnumerator GetEnumerator()
        {
            return this._readOnlyCollection.GetEnumerator();
        }

        public int IndexOf(Object value)
        {
            return this._readOnlyCollection.IndexOf(value);
        }
    }
}