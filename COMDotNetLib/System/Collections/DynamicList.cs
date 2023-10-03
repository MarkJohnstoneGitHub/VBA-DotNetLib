using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(false)]
    public class DynamicList<Type>
    {
        private List<Type> _dynamicList;


        public DynamicList() 
        {
            _dynamicList = new List<Type>();
        }

        public DynamicList(Type type)
        {
            _dynamicList = CreateFromType((dynamic)type);
        }

        public DynamicList(object type)
        {
            _dynamicList = CreateFromType((dynamic)type.GetType());
        }

        internal static List<Type> CreateFromType<Type>(Type type)
        {
            return new List<Type>();
        }

        // Properties

        public List<Type> List =>  _dynamicList;

    }
}
