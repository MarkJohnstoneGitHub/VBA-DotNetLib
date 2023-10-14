// https://learn.microsoft.com/en-us/dotnet/api/system.reflection.memberinfo?view=netframework-4.8.1

using GCollections = global::System.Collections;
using GReflection = global::System.Reflection;

namespace DotNetLib.System.Reflection
{
    public interface IMemberInfo
    {
        //GGeneric.IEnumerable<GReflection.CustomAttributeData> CustomAttributes { get; }
        GCollections.IEnumerable CustomAttributes { get; }
        Type DeclaringType { get; }

        GReflection.MemberTypes MemberType { get; }

        int MetadataToken { get; }

        GReflection.Module Module { get; }

        string Name { get; }

        Type ReflectedType { get; }

        bool Equals(object obj);

        object[] GetCustomAttributes(bool inherit);

        object[] GetCustomAttributes(Type attributeType, bool inherit);

        GCollections.IList GetCustomAttributesData();

        int GetHashCode();

        bool IsDefined(Type attributeType, bool inherit);




    }
}
