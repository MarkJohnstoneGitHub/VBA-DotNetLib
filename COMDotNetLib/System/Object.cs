// https://learn.microsoft.com/en-us/dotnet/api/system.object?view=netframework-4.8.1

using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.Extensions;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("21FD0003-0E90-41F1-BE46-BB4B59F719E9")]
    [ProgId("DotNetLib.System.Object")]
    [Description("Supports all classes in the .NET class hierarchy and provides low-level services to derived classes. This is the ultimate base class of all .NET classes; it is the root of the type hierarchy.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IObect))]
    public class Object : IObect, IWrappedObject
    {
        private object _object;
        // cache type
        //private Type _type; /??
       
        public Object()
        {
            WrappedObject = new object();
            //_type = new Type(_object.GetType());
        }

        public Object(object obj)
        {
            WrappedObject = obj; 
        }

        // Properties

        public object WrappedObject
        {
            get { return _object; }
            set
            {
                _object = value;
                //_type = new Type(value.GetType());
            }
        }

        public  new string ToString()
        { 
            return WrappedObject.ToString(); 
        }

        public new bool Equals(object obj)
        { 
            return WrappedObject.Equals(obj); 
        }

        public new int GetHashCode()
        { 
            return WrappedObject.GetHashCode();
        }

        public new Type GetType()
        {
            return new Type(_object.GetType()); ;
        }

    }
}


//// cache type //??
//private Type _type;

//public Object():base()
//        {
//    _type = new Type(base.GetType());
//}

//public new Type GetType()
//{
//    return _type;
//}