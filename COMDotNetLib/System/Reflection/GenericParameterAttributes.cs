// https://learn.microsoft.com/en-us/dotnet/api/system.reflection.genericparameterattributes?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Reflection
{
    [ComVisible(true)]
    [Guid("E8CCFEED-89C7-48E7-9062-AA6036388F13")]
    [Description("Describes the constraints on a generic type parameter of a generic type or method.\r\n\r\nThis enumeration supports a bitwise combination of its member values.")]
    //
    // Summary:
    //     Describes the constraints on a generic type parameter of a generic type or method.
    [Flags]
    //[__DynamicallyInvokable]
    public enum GenericParameterAttributes
    {
        //
        // Summary:
        //     There are no special flags.
        //[__DynamicallyInvokable]
        None = 0x0,
        //
        // Summary:
        //     Selects the combination of all variance flags. This value is the result of using
        //     logical OR to combine the following flags: System.Reflection.GenericParameterAttributes.Contravariant
        //     and System.Reflection.GenericParameterAttributes.Covariant.
        //[__DynamicallyInvokable]
        VarianceMask = 0x3,
        //
        // Summary:
        //     The generic type parameter is covariant. A covariant type parameter can appear
        //     as the result type of a method, the type of a read-only field, a declared base
        //     type, or an implemented interface.
        //[__DynamicallyInvokable]
        Covariant = 0x1,
        //
        // Summary:
        //     The generic type parameter is contravariant. A contravariant type parameter can
        //     appear as a parameter type in method signatures.
        //[__DynamicallyInvokable]
        Contravariant = 0x2,
        //
        // Summary:
        //     Selects the combination of all special constraint flags. This value is the result
        //     of using logical OR to combine the following flags: System.Reflection.GenericParameterAttributes.DefaultConstructorConstraint,
        //     System.Reflection.GenericParameterAttributes.ReferenceTypeConstraint, and System.Reflection.GenericParameterAttributes.NotNullableValueTypeConstraint.
        //[__DynamicallyInvokable]
        SpecialConstraintMask = 0x1C,
        //
        // Summary:
        //     A type can be substituted for the generic type parameter only if it is a reference
        //     type.
        //[__DynamicallyInvokable]
        ReferenceTypeConstraint = 0x4,
        //
        // Summary:
        //     A type can be substituted for the generic type parameter only if it is a value
        //     type and is not nullable.
        //[__DynamicallyInvokable]
        NotNullableValueTypeConstraint = 0x8,
        //
        // Summary:
        //     A type can be substituted for the generic type parameter only if it has a parameterless
        //     constructor.
        //[__DynamicallyInvokable]
        DefaultConstructorConstraint = 0x10
    }
}
