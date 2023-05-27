using GSystem = global::System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("FA145F11-5B72-4169-A318-BE970A28F985")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IArrayList
    {
        [Description("Initializes a new instance of the ArrayList class that is empty and has the default initial capacity.")]
        ArrayList Create();

        [Description("Initializes a new instance of the ArrayList class that is empty and has the specified initial capacity.")]
        ArrayList Create2(int capacity);

        [Description("Initializes a new instance of the ArrayList class that contains elements copied from the specified collection and that has the same initial capacity as the number of elements copied.")]
        ArrayList Create3(GSystem.Collections.ICollection c);

        //[Description("")]

        // Properties

        [Description("Gets or sets the number of elements that the ArrayList can contain.")]
        int Capacity { get; set; }

        [Description("Gets the number of elements actually contained in the ArrayList.")]
        int Count { get; }

        [Description("Gets a value indicating whether the ArrayList has a fixed size.")]
        bool IsFixedSize { get; }
    }
}
