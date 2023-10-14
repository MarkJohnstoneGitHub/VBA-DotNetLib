// https://learn.microsoft.com/en-us/dotnet/api/system.reflection.memberinfo?view=netframework-4.8.1

using GReflection = global::System.Reflection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Reflection;

namespace DotNetLib.System.Reflection
{
    public class MemberInfo : IMemberInfo
    {
        private readonly GReflection.MemberInfo _memberInfo;

        public IEnumerable CustomAttributes => throw new NotImplementedException();

        public Type DeclaringType => throw new NotImplementedException();

        public MemberTypes MemberType => throw new NotImplementedException();

        public int MetadataToken => throw new NotImplementedException();

        public Module Module => throw new NotImplementedException();

        public string Name => throw new NotImplementedException();

        public Type ReflectedType => throw new NotImplementedException();

        public object[] GetCustomAttributes(bool inherit)
        {
            throw new NotImplementedException();
        }

        public object[] GetCustomAttributes(Type attributeType, bool inherit)
        {
            throw new NotImplementedException();
        }

        public IList GetCustomAttributesData()
        {
            throw new NotImplementedException();
        }

        public bool IsDefined(Type attributeType, bool inherit)
        {
            throw new NotImplementedException();
        }
    }
}
