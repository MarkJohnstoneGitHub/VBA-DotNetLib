// https://learn.microsoft.com/en-us/dotnet/api/system.type?view=netframework-4.8.1

using GSystem = global::System;
using GType = global::System.Type;
using GReflection = global::System.Reflection;
using GInteropServices = global::System.Runtime.InteropServices;
using System;
using System.Threading.Tasks;
using DotNetLib.Extensions;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.System.Reflection;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("8E2FCAD7-C3F2-4A88-8C4E-56CF1E6E019D")]
    [ProgId("DotNetLib.System.Type")]
    [Description("Represents type declarations: class types, interface types, array types, value types, enumeration types, type parameters, generic type definitions, and open or closed constructed generic types.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IType))]
    public class Type : IType, IWrappedObject    
    {
        private GSystem.Type _type;

        private Type _baseType;
        private Type _declaringType;
        private Type[] _genericTypeArguments;
        private Type _reflectedType;
        private Type _underlyingSystemType;

        internal Type(GSystem.Type type)
        {
            _type = type;
            if (type.BaseType != null){ _baseType = new Type(type.BaseType);}
            if (type.DeclaringType != null) { _declaringType = new Type(type.DeclaringType); }
            _genericTypeArguments = WrapTypeArray(type.GenericTypeArguments);
            if (type.ReflectedType != null) { _reflectedType = new Type(type.ReflectedType); }
            //Todo: Causing Ms-Access to crash?
            //if (type.UnderlyingSystemType != null) { _underlyingSystemType = new Type(type.UnderlyingSystemType); }
           //_underlyingSystemType = new Type(type.UnderlyingSystemType);
        }

        internal GSystem.Type WrappedType => _type;

        public object WrappedObject => _type;

        public GReflection.Assembly Assembly => _type.Assembly;

        public string AssemblyQualifiedName => _type.AssemblyQualifiedName;

        public GReflection.TypeAttributes Attributes => _type.Attributes;

        public Type BaseType => _baseType;

        public virtual bool ContainsGenericParameters => _type.ContainsGenericParameters;

        public virtual GReflection.MethodBase DeclaringMethod => _type.DeclaringMethod;

        public Type DeclaringType => _declaringType;

        public static GReflection.Binder DefaultBinder => GSystem.Type.DefaultBinder;

        public string FullName => _type.FullName;

        public virtual GenericParameterAttributes GenericParameterAttributes => (GenericParameterAttributes)_type.GenericParameterAttributes;

        public virtual int GenericParameterPosition => _type.GenericParameterPosition;

        public virtual Type[] GenericTypeArguments => _genericTypeArguments;

        public  Guid GUID =>  _type.GUID;

        public bool HasElementType => _type.HasElementType;

        public bool IsAbstract => _type.IsAbstract;

        public bool IsAnsiClass => _type.IsAnsiClass;

        public bool IsAutoLayout => _type.IsAutoLayout;

        public bool IsByRef => _type.IsByRef;

        public bool IsClass => _type.IsClass;

        public bool IsCOMObject => _type.IsCOMObject;

        public virtual bool IsConstructedGenericType  => _type.IsConstructedGenericType;

        public bool IsContextful => _type.IsContextful;

        public virtual bool IsEnum => _type.IsEnum;

        public bool IsExplicitLayout  => _type.IsExplicitLayout;

        public virtual bool IsGenericParameter => _type.IsGenericParameter;

        public virtual bool IsGenericType => _type.IsGenericType;

        public virtual bool IsGenericTypeDefinition  => _type.IsGenericTypeDefinition;

        public bool IsImport => _type.IsImport;

        public bool IsInterface => _type.IsInterface;

        public bool IsLayoutSequential => _type.IsLayoutSequential;

        public bool IsMarshalByRef => _type.IsMarshalByRef;

        public bool IsNested => _type.IsNested;

        public bool IsNestedAssembly => _type.IsNestedAssembly;

        public bool IsNestedFamANDAssem => _type.IsNestedFamANDAssem;

        public bool IsNestedFamily => _type.IsNestedFamily;

        public bool IsNestedFamORAssem => _type.IsNestedFamORAssem;

        public bool IsNestedPrivate => _type.IsNestedPrivate;
        
        public bool IsNestedPublic => _type.IsNestedPublic;

        public bool IsNotPublic => _type.IsNotPublic;

        public bool IsPointer => _type.IsPointer;

        public bool IsPrimitive => _type.IsPrimitive;

        public bool IsPublic => _type.IsPublic;

        public bool IsSealed => _type.IsSealed;

        public virtual bool IsSecurityCritical => _type.IsSecurityCritical;

        public virtual bool IsSecuritySafeCritical => _type.IsSecuritySafeCritical;

        public virtual bool IsSecurityTransparent => _type.IsSecurityTransparent;

        public virtual bool IsSerializable => _type.IsSerializable;

        public bool IsSpecialName => _type.IsSpecialName;

        public bool IsUnicodeClass  => _type.IsUnicodeClass;

        public bool IsValueType => _type.IsValueType;

        public bool IsVisible => _type.IsVisible;

        public GReflection.MemberTypes MemberType => _type.MemberType;

        public GReflection.Module Module => _type.Module;

        public string Namespace => _type.Namespace;

        public Type ReflectedType => _reflectedType;

        public virtual GInteropServices.StructLayoutAttribute StructLayoutAttribute => _type.StructLayoutAttribute;

        public virtual GSystem.RuntimeTypeHandle TypeHandle =>  _type.TypeHandle;

        public GReflection.ConstructorInfo TypeInitializer => _type.TypeInitializer;

        //TODO: Causing Ms-Access to crash?
        //public Type UnderlyingSystemType => _underlyingSystemType;

        // Methods

        public virtual bool Equals(Type o)
        {
            return _type.Equals(o.WrappedObject);
        }

        public override bool Equals(object o)
        { 
            return _type.Equals(o.Unwrap()); 
        }

        public virtual Type[] FindInterfaces(GReflection.TypeFilter filter, object filterCriteria)
        {
            return WrapTypeArray(_type.FindInterfaces(filter, filterCriteria));
        }

        public virtual GReflection.MemberInfo[] FindMembers(GReflection.MemberTypes memberType, GReflection.BindingFlags bindingAttr, GReflection.MemberFilter filter, object filterCriteria)
        {
            return _type.FindMembers(memberType, bindingAttr, filter, filterCriteria);
        }

        public virtual int GetArrayRank()
        { 
            return _type.GetArrayRank(); 
        }

        public GReflection.ConstructorInfo GetConstructor(GReflection.BindingFlags bindingAttr, GReflection.Binder binder, GReflection.CallingConventions callConvention, Type[] types, GReflection.ParameterModifier[] modifiers)
        {
            return _type.GetConstructor(bindingAttr, binder, callConvention, (GSystem.Type[])types.Unwrap(), modifiers);
        }

        public GReflection.ConstructorInfo[] GetConstructors()
        {
            return _type.GetConstructors();
        }

        public virtual GReflection.MemberInfo[] GetDefaultMembers()
        { 
            return _type.GetDefaultMembers();
        }

        public Type GetElementType()
        {
            return new Type(_type.GetElementType());
        }

        public virtual string GetEnumName(object value)
        { 
            return _type.GetEnumName(value);
        }

        public virtual string[] GetEnumNames()
        { 
            return _type.GetEnumNames(); 
        }

        public virtual Type GetEnumUnderlyingType()
        {
            return new Type(_type.GetEnumUnderlyingType());
        }

        public virtual Array GetEnumValues()
        {
            return new Array(_type.GetEnumValues());
        }

        public GReflection.EventInfo GetEvent(string name)
        { 
            return _type.GetEvent(name);
        }

        public GReflection.EventInfo GetEvent(string name, GReflection.BindingFlags bindingAttr)
        {
            return _type.GetEvent(name, bindingAttr);
        }

        public virtual GReflection.EventInfo[] GetEvents()
        {
            return _type.GetEvents();
        }

        //public abstract GReflection.EventInfo[] GetEvents(GReflection.BindingFlags bindingAttr);

        public GReflection.MemberInfo[] GetMembers()
        { 
            return _type.GetMembers(); 
        }

        public override int GetHashCode()
        { 
            return _type.GetHashCode(); 
        }

        //Todo : Issue throwing exception when throwOnError is false? https://github.com/dotnet/runtime/issues/12376
        public static Type GetType(string typeName)
        {
            return new Type(GType.GetType(typeName));
        }

        public static Type GetType(string typeName, bool throwOnError)
        {
            return new Type(GType.GetType(typeName, throwOnError));
        }

        public static Type GetType(string typeName, bool throwOnError, bool ignoreCase)
        {
            return new Type(GType.GetType(typeName, throwOnError, ignoreCase));
        }

        public static TypeCode GetTypeCode(Type type)
        {
            return GType.GetTypeCode(type.WrappedType);
        }

        public virtual Type MakeArrayType()
        {
            return new Type(_type.MakeArrayType());
        }

        public virtual Type MakeArrayType(int rank)
        {
            return new Type(_type.MakeArrayType(rank));
        }

        public override string ToString()
        { 
            return _type.ToString(); 
        }


        // Wrap Array of GSystem.Type[] to DotNetLib.System.Type
        internal Type[] WrapTypeArray(GSystem.Type[] types)
        {
            if (types == null)
            {
                return null;
            }
            Type[] wrappedTypes = new Type[types.Length];
            for (int index = 0; index < wrappedTypes.Length; index++)
            {
                wrappedTypes[index] = new Type(types[index]);
            }
            return wrappedTypes;
        }

        public GReflection.FieldInfo GetField(string name)
        {
            return _type.GetField(name);
        }

        public GReflection.FieldInfo[] GetFields()
        {
            return _type.GetFields();
        }
    }
}
