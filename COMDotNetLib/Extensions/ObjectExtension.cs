namespace DotNetLib.Extensions
{
    public static class ObjectExtension
    {
        public static object Unwrap(this object obj) 
        {
            if (obj is IWrappedObject unwrappedObj)
            {
                return unwrappedObj.WrappedObject;
            }
            return obj; //If not a wrapped object return orignal object
        }

        public static object[] Unwrap(this object[] objs)
        {
            if (objs == null)
                return null;

            object[] unwrapObjs = new object[objs.Length];
            for (int index = 0; index < unwrapObjs.Length; index++)
            {
                unwrapObjs[index] = objs[index].Unwrap();
            }
            return unwrapObjs;
        }

        public static bool IsWrapped(this object obj)
        {
            return obj is IWrappedObject;
        }

    }
}
