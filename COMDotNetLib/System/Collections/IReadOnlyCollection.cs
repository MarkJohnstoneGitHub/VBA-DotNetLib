using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("2BFD4F86-013C-49D1-A5A5-AD0EEBD10E30")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IReadOnlyCollection
    {
        [Description("Initializes a new instance of the ReadOnlyCollection<T> class that is a read-only wrapper around the specified list.")]
        ReadOnlyCollection Create(IList list);

        [Description("Gets the number of elements contained in the ReadOnlyCollection<T> instance.")]
        int Count { get; }

        [Description("Gets the element at the specified index.")]
        object this[int index] { get; }

        [Description("Determines whether an element is in the ReadOnlyCollection<T>.")]
        bool Contains(object Index);

        [Description("Copies the entire ReadOnlyCollection<T> to a compatible one-dimensional Array, starting at the specified index of the target array.")]
        void CopyTo(object[] array, int index);

        [Description("Returns an enumerator that iterates through the ReadOnlyCollection<T>.")]
        GSystem.Collections.IEnumerator GetEnumerator();

        [Description("Searches for the specified object and returns the zero-based index of the first occurrence within the entire ReadOnlyCollection<T>.")]
        int IndexOf(object value);
    }
}
