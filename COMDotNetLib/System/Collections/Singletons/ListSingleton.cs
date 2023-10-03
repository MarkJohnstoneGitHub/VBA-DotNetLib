using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a strongly typed list of objects that can be accessed by index. Provides methods to search, sort, and manipulate lists.")]
    [Guid("ABA0418D-78E1-4220-AACD-FAE3238CA4EF")]
    [ProgId("DotNetLib.System.Collections.ListSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IListSingleton))]
    public class ListSingleton : IListSingleton
    {
        public ListSingleton() { }

        public List Create()
        {
            return new List();
        }

        public List Create2(object type, int capacity = 0)
        {
            return new List(type, capacity);
        }

    }
}
