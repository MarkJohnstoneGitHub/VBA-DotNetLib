// https://referencesource.microsoft.com/#mscorlib/system/security/accesscontrol/enums.cs

using System;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Security.AccessControl
{
    [ComVisible(true)]
    [Flags]
    public enum InheritanceFlags
    {
        None = 0x00,
        ContainerInherit = 0x01,
        ObjectInherit = 0x02,
    }

    [ComVisible(true)]
    [Flags]
    public enum PropagationFlags
    {
        None = 0x00,
        NoPropagateInherit = 0x01,
        InheritOnly = 0x02,
    }

    [ComVisible(true)]
    [Flags]
    public enum AuditFlags
    {
        None = 0x00,
        Success = 0x01,
        Failure = 0x02,
    }


    [ComVisible(true)]
    [Flags]
    public enum SecurityInfos
    {
        Owner = 0x00000001,
        Group = 0x00000002,
        DiscretionaryAcl = 0x00000004,
        SystemAcl = 0x00000008,

    }


    [ComVisible(true)]
    public enum ResourceType
    {
        Unknown = 0x00,
        FileObject = 0x01,
        Service = 0x02,
        Printer = 0x03,
        RegistryKey = 0x04,
        LMShare = 0x05,
        KernelObject = 0x06,
        WindowObject = 0x07,
        DSObject = 0x08,
        DSObjectAll = 0x09,
        ProviderDefined = 0x0A,
        WmiGuidObject = 0x0B,
        RegistryWow6432Key = 0x0C,
    }

    [ComVisible(true)]
    [Guid("71F3AD31-8BF3-437B-86E5-7CD3CBE5C523")]

    //
    // Summary:
    //     Specifies which sections of a security descriptor to save or load.
    [Flags]
    public enum AccessControlSections
    {
        //
        // Summary:
        //     No sections.
        None = 0x0,
        //
        // Summary:
        //     The system access control list (SACL).
        Audit = 0x1,
        //
        // Summary:
        //     The discretionary access control list (DACL).
        Access = 0x2,
        //
        // Summary:
        //     The owner.
        Owner = 0x4,
        //
        // Summary:
        //     The primary group.
        Group = 0x8,
        //
        // Summary:
        //     The entire security descriptor.
        All = 0xF
    }

    [ComVisible(true)]
    [Flags]
    public enum AccessControlActions
    {
#if FEATURE_MACL
        None = 0,
        View = 1,
        Change = 2
#else
        None = 0
#endif
    }
}
