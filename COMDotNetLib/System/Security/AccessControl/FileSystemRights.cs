﻿// https://learn.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.filesystemrights?view=netframework-4.8.1
// https://referencesource.microsoft.com/#mscorlib/system/security/accesscontrol/filesecurity.cs,30

using System;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Security.AccessControl
{
    [ComVisible(true)]
    [Guid("165D1915-BB88-4675-BDD5-3CE7F86E709D")]

    // Constants from from winnt.h - search for FILE_WRITE_DATA, etc.
    [Flags]
    public enum FileSystemRights
    {
        // No None field - An ACE with the value 0 cannot grant nor deny.
        ReadData = 0x000001,
        ListDirectory = ReadData,     // For directories
        WriteData = 0x000002,
        CreateFiles = WriteData,    // For directories
        AppendData = 0x000004,
        CreateDirectories = AppendData,   // For directories
        ReadExtendedAttributes = 0x000008,
        WriteExtendedAttributes = 0x000010,
        ExecuteFile = 0x000020,     // For files
        Traverse = ExecuteFile,  // For directories
        // DeleteSubdirectoriesAndFiles only makes sense on directories, but 
        // the shell explicitly sets it for files in its UI.  So we'll include 
        // it in FullControl.
        DeleteSubdirectoriesAndFiles = 0x000040,
        ReadAttributes = 0x000080,
        WriteAttributes = 0x000100,
        Delete = 0x010000,
        ReadPermissions = 0x020000,
        ChangePermissions = 0x040000,
        TakeOwnership = 0x080000,
        // From the Core File Services team, CreateFile always requires 
        // SYNCHRONIZE access.  Very tricksy, CreateFile is.
        Synchronize = 0x100000,  // Can we wait on the handle?
        FullControl = 0x1F01FF,

        // These map to what Explorer sets, and are what most users want.
        // However, an ACL editor will also want to set the Synchronize
        // bit when allowing access, and exclude the synchronize bit when
        // denying access.
        Read = ReadData | ReadExtendedAttributes | ReadAttributes | ReadPermissions,
        ReadAndExecute = Read | ExecuteFile,
        Write = WriteData | AppendData | WriteExtendedAttributes | WriteAttributes,
        Modify = ReadAndExecute | Write | Delete,
    }
}
