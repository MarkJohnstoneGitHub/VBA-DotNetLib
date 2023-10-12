// https://learn.microsoft.com/en-us/dotnet/api/system.type?view=netframework-4.8.1

using GSystem = global::System;
using GReflection = global::System.Reflection;
using GInteropServices = global::System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System;
using DotNetLib.System.Reflection;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("4108540D-1DA8-4411-BC25-F641145973B5")]
    [Description("Represents type declarations: class types, interface types, array types, value types, enumeration types, type parameters, generic type definitions, and open or closed constructed generic types.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IType
    {
        GReflection.Assembly Assembly 
        {
            [Description("Gets the Assembly in which the type is declared. For generic types, gets the Assembly in which the generic type is defined.")]
            get;
        }

        string AssemblyQualifiedName 
        {
            [Description("Gets the assembly-qualified name of the type, which includes the name of the assembly from which this Type object was loaded.")]
            get;
        }

        GReflection.TypeAttributes Attributes 
        {
            [Description("Gets the attributes associated with the Type.")]
            get; 
        }

        Type BaseType 
        {
            [Description("Gets the type from which the current Type directly inherits.")]
            get;
        }

        bool ContainsGenericParameters 
        {
            [Description("Gets a value indicating whether the current Type object has type parameters that have not been replaced by specific types.")]
            get;
        }

        GReflection.MethodBase DeclaringMethod 
        {
            [Description("Gets a MethodBase that represents the declaring method, if the current Type represents a type parameter of a generic method.")]
            get;
        }

        Type DeclaringType 
        {
            [Description("Gets the type that declares the current nested type or generic type parameter.")]
            get;
        }

        string FullName 
        {
            [Description("Gets the fully qualified name of the type, including its namespace but not its assembly.")]
            get;
        }

        GenericParameterAttributes GenericParameterAttributes
        {
            [Description("Gets a combination of GenericParameterAttributes flags that describe the covariance and special constraints of the current generic type parameter.")]
            get;
        }

        int GenericParameterPosition 
        {
            [Description("Gets the position of the type parameter in the type parameter list of the generic type or method that declared the parameter, when the Type object represents a type parameter of a generic type or a generic method.")]
            get;
        }

        Type[] GenericTypeArguments 
        {
            [Description("Gets an array of the generic type arguments for this type.")]
            get;
        }

        GSystem.Guid GUID 
        {
            [Description("Gets the GUID associated with the Type.")]
            get;
        }

        bool HasElementType 
        {
            [Description("Gets a value indicating whether the current Type encompasses or refers to another type; that is, whether the current Type is an array, a pointer, or is passed by reference.")]
            get;
        }

        bool IsAbstract 
        {
            [Description("Gets a value indicating whether the Type is abstract and must be overridden.")]
            get;
        }

        bool IsAnsiClass 
        {
            [Description("Gets a value indicating whether the string format attribute AnsiClass is selected for the Type.")]
            get;
        }

        bool IsAutoLayout 
        {
            [Description("Gets a value indicating whether the fields of the current type are laid out automatically by the common language runtime.")]
            get;
        }

        bool IsByRef 
        {
            [Description("Gets a value indicating whether the Type is passed by reference.")]
            get;
        }

        bool IsClass 
        {
            [Description("Gets a value indicating whether the Type is a class or a delegate; that is, not a value type or interface.")]
            get;
        }

        bool IsCOMObject 
        {
            [Description("Gets a value indicating whether the Type is a COM object.")]
            get;
        }

        bool IsConstructedGenericType 
        {
            [Description("Gets a value that indicates whether this object represents a constructed generic type. You can create instances of a constructed generic type.")]
            get;
        }

        bool IsContextful 
        {
            [Description("Gets a value indicating whether the Type can be hosted in a context.")]
            get;
        }

        bool IsEnum 
        {
            [Description("Gets a value indicating whether the current Type represents an enumeration.")]
            get;
        }

        bool IsExplicitLayout 
        {
            [Description("Gets a value indicating whether the fields of the current type are laid out at explicitly specified offsets.")]
            get;
        }

        bool IsGenericParameter 
        {
            [Description("Gets a value indicating whether the current Type represents a type parameter in the definition of a generic type or method.")]
            get;
        }

        bool IsGenericType 
        {
            [Description("Gets a value indicating whether the current type is a generic type.")]
            get;
        }

        bool IsGenericTypeDefinition 
        {
            [Description("Gets a value indicating whether the current Type represents a generic type definition, from which other generic types can be constructed.")]
            get;
        }

        bool IsImport 
        {
            [Description("Gets a value indicating whether the Type has a ComImportAttribute attribute applied, indicating that it was imported from a COM type library.")]
            get;
        }

        bool IsInterface 
        {
            [Description("Gets a value indicating whether the Type is an interface; that is, not a class or a value type.")]
            get;
        }

        bool IsLayoutSequential 
        {
            [Description("Gets a value indicating whether the fields of the current type are laid out sequentially, in the order that they were defined or emitted to the metadata.")]
            get;
        }

        bool IsMarshalByRef 
        {
            [Description("Gets a value indicating whether the Type is marshaled by reference.")]
            get;
        }

        bool IsNested 
        {
            [Description("Gets a value indicating whether the current Type object represents a type whose definition is nested inside the definition of another type.")]
            get;
        }

        bool IsNestedAssembly 
        {
            [Description("Gets a value indicating whether the Type is nested and visible only within its own assembly.")]
            get;
        }

        bool IsNestedFamANDAssem 
        {
            [Description("Gets a value indicating whether the Type is nested and visible only to classes that belong to both its own family and its own assembly.")]
            get;
        }

        bool IsNestedFamily 
        {
            [Description("Gets a value indicating whether the Type is nested and visible only within its own family.")]
            get;
        }

        bool IsNestedFamORAssem 
        {
            [Description("Gets a value indicating whether the Type is nested and visible only to classes that belong to either its own family or to its own assembly.")]
            get;
        }

        bool IsNestedPrivate 
        {
            [Description("Gets a value indicating whether the Type is nested and declared private.")]
            get;
        }

        bool IsNestedPublic 
        {
            [Description("Gets a value indicating whether a class is nested and declared public.")]
            get;
        }

        bool IsNotPublic 
        {
            [Description("Gets a value indicating whether the Type is not declared public.")]
            get;
        }

        bool IsPointer 
        {
            [Description("Gets a value indicating whether the Type is a pointer.")]
            get;
        }

        bool IsPrimitive 
        {
            [Description("Gets a value indicating whether the Type is one of the primitive types.")]
            get;
        }

        bool IsPublic 
        {
            [Description("Gets a value indicating whether the Type is declared public.")]
            get;
        }

        bool IsSealed 
        {
            [Description("Gets a value indicating whether the Type is declared sealed.")]
            get;
        }

        bool IsSecurityCritical 
        {
            [Description("Gets a value that indicates whether the current type is security-critical or security-safe-critical at the current trust level, and therefore can perform critical operations.")]
            get;
        }

        bool IsSecuritySafeCritical 
        {
            [Description("Gets a value that indicates whether the current type is security-safe-critical at the current trust level; that is, whether it can perform critical operations and can be accessed by transparent code.")]
            get;
        }

        bool IsSecurityTransparent 
        {
            [Description("Gets a value that indicates whether the current type is transparent at the current trust level, and therefore cannot perform critical operations.")]
            get;
        }

        bool IsSerializable 
        {
            [Description("Gets a value indicating whether the Type is binary serializable.")]
            get;
        }

        bool IsSpecialName 
        {
            [Description("Gets a value indicating whether the type has a name that requires special handling.")]
            get;
        }

        bool IsUnicodeClass 
        {
            [Description("Gets a value indicating whether the string format attribute UnicodeClass is selected for the Type.")]
            get;
        }

        bool IsValueType 
        {
            [Description("Gets a value indicating whether the Type is a value type.")]
            get;
        }

        bool IsVisible 
        {
            [Description("Gets a value indicating whether the Type can be accessed by code outside the assembly.")]
            get;
        }

        GReflection.MemberTypes MemberType 
        {
            [Description("Gets a MemberTypes value indicating that this member is a type or a nested type.")]
            get;
        }

        GReflection.Module Module 
        {
            [Description("Gets the module (the DLL) in which the current Type is defined.")]
            get;
        }

        string Namespace 
        {
            [Description("Gets the namespace of the Type.")]
            get;
        }

        Type ReflectedType 
        {
            [Description("Gets the class object that was used to obtain this member.")]
            get;
        }

        GInteropServices.StructLayoutAttribute StructLayoutAttribute 
        {
            [Description("Gets a StructLayoutAttribute that describes the layout of the current type.")]
            get;
        }

        GSystem.RuntimeTypeHandle TypeHandle 
        {
            [Description("Gets the handle for the current Type.")]
            get;
        }

        GReflection.ConstructorInfo TypeInitializer 
        {
            [Description("Gets the initializer for the type.")]
            get;
        }

        //Type UnderlyingSystemType
        //{
        //    [Description("Indicates the type provided by the common language runtime that represents this type.")]
        //    get;
        //}

        // Methods

        [Description("Determines if the underlying system type of the current Type is the same as the underlying system type of the specified Type.")]
        bool Equals(Type o);

        [Description("Returns an array of Type objects representing a filtered list of interfaces implemented or inherited by the current Type.")]
        Type[] FindInterfaces(GReflection.TypeFilter filter, object filterCriteria);

        [Description("Returns a filtered array of MemberInfo objects of the specified member type.")]
        GReflection.MemberInfo[] FindMembers(GReflection.MemberTypes memberType, GReflection.BindingFlags bindingAttr, GReflection.MemberFilter filter, object filterCriteria);

        [Description("Gets the number of dimensions in an array.")]
        int GetArrayRank();

        [Description("Searches for a constructor whose parameters match the specified argument types and modifiers, using the specified binding constraints and the specified calling convention.")]
        GReflection.ConstructorInfo GetConstructor(GReflection.BindingFlags bindingAttr, GReflection.Binder binder, GReflection.CallingConventions callConvention, Type[] types, GReflection.ParameterModifier[] modifiers);

        [Description("Returns all the public constructors defined for the current Type.")]
        GReflection.ConstructorInfo[] GetConstructors();

        [Description("Searches for the members defined for the current Type whose DefaultMemberAttribute is set.")] 
        GReflection.MemberInfo[] GetDefaultMembers();

        [Description("When overridden in a derived class, returns the Type of the object encompassed or referred to by the current array, pointer or reference type.")]
        Type GetElementType();

        [Description("Returns the name of the constant that has the specified value, for the current enumeration type.")]
        string GetEnumName(object value);

        [Description("Returns the names of the members of the current enumeration type.")]
        string[] GetEnumNames();

        [Description("Returns the underlying type of the current enumeration type.")]
        Type GetEnumUnderlyingType();

        [Description("Returns an array of the values of the constants in the current enumeration type.")]
        Array GetEnumValues();

        [Description("Returns the EventInfo object representing the specified public event.")]
        GReflection.EventInfo GetEvent(string name);

        [Description("When overridden in a derived class, returns the EventInfo object representing the specified event, using the specified binding constraints.")]
        GReflection.EventInfo GetEvent(string name, GReflection.BindingFlags bindingAttr);

        [Description("Returns all the public events that are declared or inherited by the current Type.")]
        GReflection.EventInfo[] GetEvents();

        [Description("Searches for the public field with the specified name.")]
        GReflection.FieldInfo GetField(string name);

        [Description("Returns all the public fields of the current Type.")]
        GReflection.FieldInfo[] GetFields();

        [Description("Returns the hash code for this instance.")]
        int GetHashCode();

        [Description("Returns all the public members of the current Type.")]
        GReflection.MemberInfo[] GetMembers();

        [Description("Returns a Type object representing a one-dimensional array of the current type, with a lower bound of zero.")]
        Type MakeArrayType();

        [Description("Returns a Type object representing an array of the current type, with the specified number of dimensions.")]
        Type MakeArrayType(int rank);


        [Description("Returns a String representing the name of the current Type.")]
        string ToString();


        // [Description("")]



    }
}
