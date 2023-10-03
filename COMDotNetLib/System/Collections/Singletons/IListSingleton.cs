using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("B8DE48EC-9EFF-4417-800D-090BA557DB77")]
    [Description("")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IListSingleton
    {
        [Description("Initializes a new instance of the List<object> class that is empty and has the default initial capacity.")]
        List Create();

        [Description("Initializes a new instance of the List class of the type provided that is empty and has the default or specified initial capacity.")]
        List Create2(object type, int capacity  = 0);

    }
}
