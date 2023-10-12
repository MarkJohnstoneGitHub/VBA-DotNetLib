// https://learn.microsoft.com/en-us/dotnet/api/system.object?view=netframework-4.8.1

using DotNetLib.Extensions;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("092463AE-6DFE-40FF-BA1C-2872E65F295D")]
    [ProgId("DotNetLib.System.ObjectSingleton")]
    [Description("")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IObectSingleton))]
    public class ObjectSingleton : IObectSingleton
    {
        //public ObjectSingleton() { }

        public Object Create(object obj = null)
        { 
            if (obj == null)
            {
                return  new Object();
            }
            return new Object(obj); 
        }

        public new bool Equals(object objA, object objB)
        {  
            return Object.Equals(objA.Unwrap(), objB.Unwrap()); 
        }

        public new bool ReferenceEquals(object objA, object objB)
        {
            return Object.ReferenceEquals(objA.Unwrap(), objB.Unwrap());
        }

    }
}
