using System;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("39863A22-78AC-4203-84E9-AB40BA29F394")]

    //
    // Summary:
    //     Specifies options to use for getting the path to a special folder.
    public enum SpecialFolderOption
    {
        //
        // Summary:
        //     The path to the folder is verified. If the folder exists, the path is returned.
        //     If the folder does not exist, an empty string is returned. This is the default
        //     behavior.
        None = 0,
        //
        // Summary:
        //     The path to the folder is created if it does not already exist.
        Create = 0x8000,
        //
        // Summary:
        //     The path to the folder is returned without verifying whether the path exists.
        //     If the folder is located on a network, specifying this option can reduce lag
        //     time.
        DoNotVerify = 0x4000
    }
}
