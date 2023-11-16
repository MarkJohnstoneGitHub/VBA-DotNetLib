// https://learn.microsoft.com/en-us/dotnet/api/system.io.fileinfo?view=netframework-4.8.1

using GSerialization = global::System.Runtime.Serialization;
using GIO = global::System.IO;
using GAccessControl = global::System.Security.AccessControl;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.IO;
using DotNetLib.System.Security.AccessControl;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("1EF84DBA-6ACD-4C81-9483-DFA41CBEED47")]
    [Description("Provides properties and instance methods for the creation, copying, deletion, moving, and opening of files, and aids in the creation of FileStream objects. This class cannot be inherited.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IFileInfo
    {
        DirectoryInfo Directory
        {
            [Description("Gets an instance of the parent directory.")]
            get;
        }

        string DirectoryName
        {
            [Description("Gets a string representing the directory's full path.")]
            get;
        }

        bool Exists 
        {
            [Description("Gets a value indicating whether a file exists.")]
            get;
        }

        bool IsReadOnly
        {
            [Description("Gets or sets a value that determines if the current file is read only.")]
            get;
            [Description("Gets or sets a value that determines if the current file is read only.")]
            set;
        }

        long Length
        {
            [Description("Gets the size, in bytes, of the current file.")]
            get;
        }

        string Name 
        {
            [Description("Gets the name of the file.")]
            get;
        }

        GIO.FileAttributes Attributes
        {
            [Description("Gets or sets the attributes for the current file or directory.")]
            get;
            [Description("Gets or sets the attributes for the current file or directory.")]
            set;
        }

        DateTime CreationTime
        {
            [Description("Gets or sets the creation time of the current file or directory.")]
            get;
            [Description("Gets or sets the creation time of the current file or directory.")]
            set;
        }

        DateTime CreationTimeUtc
        {
            [Description("Gets or sets the creation time, in coordinated universal time (UTC), of the current file or directory.")]
            get;
            [Description("Gets or sets the creation time, in coordinated universal time (UTC), of the current file or directory.")]
            set;
        }

        //bool Exists
        //{
        //    [Description("Gets a value indicating whether the file or directory exists.")]
        //    get;
        //}

        string Extension
        {
            [Description("Gets the extension part of the file name, including the leading dot . even if it is the entire file name, or an empty string if no extension is present.")]
            get;
        }

        string FullName
        {
            [Description("Gets the full path of the directory or file.")]
            get;
        }

        DateTime LastAccessTime
        {
            [Description("Gets or sets the time the current file or directory was last accessed.")]
            get;
            [Description("Gets or sets the time the current file or directory was last accessed.")]
            set;
        }

        DateTime LastAccessTimeUtc
        {
            [Description("Gets or sets the time, in coordinated universal time (UTC), that the current file or directory was last accessed.")]
            get;
            [Description("Gets or sets the time, in coordinated universal time (UTC), that the current file or directory was last accessed.")]
            set;
        }

        DateTime LastWriteTime
        {
            [Description("Gets or sets the time when the current file or directory was last written to.")]
            get;
            [Description("Gets or sets the time when the current file or directory was last written to.")]
            set;
        }

        DateTime LastWriteTimeUtc
        {
            [Description("Gets or sets the time, in coordinated universal time (UTC), when the current file or directory was last written to.")]
            get;
            [Description("Gets or sets the time, in coordinated universal time (UTC), when the current file or directory was last written to.")]
            set;
        }

        //string Name
        //{
        //    [Description("For files, gets the name of the file. For directories, gets the name of the last directory in the hierarchy if a hierarchy exists. Otherwise, the Name property gets the name of the directory.")]
        //    get;
        //}

        // Methods

        [Description("Creates a StreamWriter that appends text to the file represented by this instance of the FileInfo.")]
        GIO.StreamWriter AppendText();

        [Description("Copies an existing file to a new file, disallowing the overwriting of an existing file.")]
        FileInfo CopyTo(string destFileName);

        [Description("Copies an existing file to a new file, allowing the overwriting of an existing file.")]
        FileInfo CopyTo(string destFileName, bool overwrite);

        [Description("Creates a file.")]
        GIO.FileStream Create();

        [Description("Creates a StreamWriter that writes a new text file.")]
        GIO.StreamWriter CreateText();

        [Description("Decrypts a file that was encrypted by the current account using the Encrypt() method.")]
        void Decrypt();

        [Description("Permanently deletes a file.")]
        void Delete();

        [Description("Encrypts a file so that only the account used to encrypt the file can decrypt it.")]
        void Encrypt();

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("Gets a FileSecurity object that encapsulates the access control list (ACL) entries for the file described by the current FileInfo object.")]
        GAccessControl.FileSecurity GetAccessControl();

        [Description("Gets a FileSecurity object that encapsulates the specified type of access control list (ACL) entries for the file described by the current FileInfo object.")]
        GAccessControl.FileSecurity GetAccessControl(AccessControlSections includeSections);

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        int GetHashCode();

        [Description("Sets the SerializationInfo object with the file name and additional exception information.\r\n\r\n(Inherited from FileSystemInfo)")]
        void GetObjectData(GSerialization.SerializationInfo info, GSerialization.StreamingContext context);

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Moves a specified file to a new location, providing the option to specify a new file name.")]
        void MoveTo(string destFileName);

        [Description("Opens a file in the specified mode.")]
        GIO.FileStream Open(GIO.FileMode mode);

        [Description("Opens a file in the specified mode with read, write, or read/write access.")]
        GIO.FileStream Open(GIO.FileMode mode, GIO.FileAccess access);

        [Description("Opens a file in the specified mode with read, write, or read/write access and the specified sharing option.")]
        GIO.FileStream Open(GIO.FileMode mode, GIO.FileAccess access, GIO.FileShare share);

        [Description("Creates a read-only FileStream.")]
        GIO.FileStream OpenRead();

        [Description("Creates a StreamReader with UTF8 encoding that reads from an existing text file.")]
        GIO.StreamReader OpenText();

        [Description("Creates a write-only FileStream.")]
        GIO.FileStream OpenWrite();

        [Description("Refreshes the state of the object.\r\n\r\n(Inherited from FileSystemInfo)")]
        void Refresh();

        [Description("Replaces the contents of a specified file with the file described by the current FileInfo object, deleting the original file, and creating a backup of the replaced file.")]
        FileInfo Replace(string destinationFileName, string destinationBackupFileName);

        [Description("Replaces the contents of a specified file with the file described by the current FileInfo object, deleting the original file, and creating a backup of the replaced file. Also specifies whether to ignore merge errors.")]
        FileInfo Replace(string destinationFileName, string destinationBackupFileName, bool ignoreMetadataErrors);

        [Description("Applies access control list (ACL) entries described by a FileSecurity object to the file described by the current FileInfo object.")]
        void SetAccessControl(GAccessControl.FileSecurity fileSecurity);

        [Description("Returns the original path that was passed to the FileInfo constructor. Use the FullName or Name property for the full path or file name.")]
        string ToString();

    }
}
