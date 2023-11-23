// https://learn.microsoft.com/en-us/dotnet/api/system.io.file?view=netframework-4.8.1

using GText = global::System.Text;
using GCollections = global::System.Collections;
using GIO = global::System.IO;
using GAccessControl = global::System.Security.AccessControl;
using DotNetLib.System.Security.AccessControl;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.System.Text;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("9DA890DB-1B5D-41B3-BA36-6FE9E8D3C1C5")]
    [Description("Provides static methods for the creation, copying, deletion, moving, and opening of a single file, and aids in the creation of FileStream objects.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IFileSingleton
    {
        [Description("Appends lines to a file, and then closes the file. If the specified file does not exist, this method creates a file, writes the specified lines to the file, and then closes the file.")]
        void AppendAllLines(string path, GCollections.IEnumerable contents);


        [Description("Appends lines to a file by using a specified encoding, and then closes the file. If the specified file does not exist, this method creates a file, writes the specified lines to the file, and then closes the file.")]
        void AppendAllLines(string path, GCollections.IEnumerable contents, Encoding encoding);

        [Description("Opens a file, appends the specified string to the file, and then closes the file. If the file does not exist, this method creates a file, writes the specified string to the file, then closes the file.")]
        void AppendAllText(string path, string contents);

        [Description("Appends the specified string to the file using the specified encoding, creating the file if it does not already exist.")]
        void AppendAllText(string path, string contents, Encoding encoding);

        [Description("Creates a StreamWriter that appends UTF-8 encoded text to an existing file, or to a new file if the specified file does not exist.")]
        StreamWriter AppendText(string path);

        [Description("Copies an existing file to a new file. Overwriting a file of the same name is not allowed.")]
        void Copy(string sourceFileName, string destFileName);

        [Description("Copies an existing file to a new file. Overwriting a file of the same name is allowed.")]
        void Copy(string sourceFileName, string destFileName, bool overwrite);

        [Description("Creates or overwrites a file in the specified path.")]
        GIO.FileStream Create(string path);

        [Description("Creates or overwrites a file in the specified path, specifying a buffer size.")]
        GIO.FileStream Create(string path, int bufferSize);

        [Description("Creates or overwrites a file in the specified path, specifying a buffer size and options that describe how to create or overwrite the file.")]
        GIO.FileStream Create(string path, int bufferSize, GIO.FileOptions options);

        [Description("Creates or overwrites a file in the specified path, specifying a buffer size, options that describe how to create or overwrite the file, and a value that determines the access control and audit security for the file.")]
        GIO.FileStream Create(string path, int bufferSize, GIO.FileOptions options, GAccessControl.FileSecurity fileSecurity);

        [Description("Creates or opens a file for writing UTF-8 encoded text. If the file already exists, its contents are overwritten.")]
        StreamWriter CreateText(string path);

        [Description("Decrypts a file that was encrypted by the current account using the Encrypt(String) method.")]
        void Decrypt(string path);

        [Description("Deletes the specified file.")]
        void Delete(string path);

        [Description("Encrypts a file so that only the account used to encrypt the file can decrypt it.")]
        void Encrypt(string path);

        [Description("Determines whether the specified file exists.")]
        bool Exists(string path);

        [Description("Gets a FileSecurity object that encapsulates the access control list (ACL) entries for a specified file.")]
        GAccessControl.FileSecurity GetAccessControl(string path);

        [Description("Gets a FileSecurity object that encapsulates the specified type of access control list (ACL) entries for a particular file.")]
        GAccessControl.FileSecurity GetAccessControl(string path, AccessControlSections includeSections);

        [Description("Gets the FileAttributes of the file on the path.")]
        GIO.FileAttributes GetAttributes(string path);

        [Description("Returns the creation date and time of the specified file or directory.")]
        DateTime GetCreationTime(string path);

        [Description("Returns the creation date and time, in Coordinated Universal Time (UTC), of the specified file or directory.")]
        DateTime GetCreationTimeUtc(string path);

        [Description("Returns the date and time the specified file or directory was last accessed.")]
        DateTime GetLastAccessTime(string path);

        [Description("Returns the date and time, in Coordinated Universal Time (UTC), that the specified file or directory was last accessed.")]
        DateTime GetLastAccessTimeUtc(string path);

        [Description("Returns the date and time the specified file or directory was last written to.")]
        DateTime GetLastWriteTime(string path);

        [Description("Returns the date and time, in Coordinated Universal Time (UTC), that the specified file or directory was last written to.")]
        DateTime GetLastWriteTimeUtc(string path);

        [Description("Moves a specified file to a new location, providing the option to specify a new file name.")]
        void Move(string sourceFileName, string destFileName);

        [Description("Opens a FileStream on the specified path with read/write access with no sharing.")]
        GIO.FileStream Open(string path, GIO.FileMode mode);

        [Description("Opens a FileStream on the specified path, with the specified mode and access with no sharing.")]
        GIO.FileStream Open(string path, GIO.FileMode mode, GIO.FileAccess access);

        [Description("Opens a FileStream on the specified path, having the specified mode with read, write, or read/write access and the specified sharing option.")]
        GIO.FileStream Open(string path, GIO.FileMode mode, GIO.FileAccess access, GIO.FileShare share);

        [Description("Opens an existing file for reading.")]
        GIO.FileStream OpenRead(string path);

        [Description("Opens an existing UTF-8 encoded text file for reading.")]
        GIO.StreamReader OpenText(string path);

        [Description("Opens an existing file or creates a new file for writing.")]
        GIO.FileStream OpenWrite(string path);

        [Description("Opens a binary file, reads the contents of the file into a byte array, and then closes the file.")]
        byte[] ReadAllBytes(string path);

        [Description("Opens a text file, reads all lines of the file, and then closes the file.")]
        string[] ReadAllLines(string path);

        [Description("Opens a file, reads all lines of the file with the specified encoding, and then closes the file.")]
        string[] ReadAllLines(string path, Encoding encoding);

        [Description("Opens a text file, reads all the text in the file, and then closes the file.")]
        string ReadAllText(string path);

        [Description("Reads the lines of a file.")]
        GCollections.IEnumerable ReadLines(string path);

        [Description("Read the lines of a file that has a specified encoding.")]
        GCollections.IEnumerable ReadLines(string path, Encoding encoding);

        [Description("Replaces the contents of a specified file with the contents of another file, deleting the original file, and creating a backup of the replaced file.")]
        void Replace(string sourceFileName, string destinationFileName, string destinationBackupFileName);

        [Description("Replaces the contents of a specified file with the contents of another file, deleting the original file, and creating a backup of the replaced file and optionally ignores merge errors.")]
        void Replace(string sourceFileName, string destinationFileName, string destinationBackupFileName, bool ignoreMetadataErrors);

        [Description("Applies access control list (ACL) entries described by a FileSecurity object to the specified file.")]
        void SetAccessControl(string path, GAccessControl.FileSecurity fileSecurity);

        [Description("Sets the specified FileAttributes of the file on the specified path.")]
        void SetAttributes(string path, GIO.FileAttributes fileAttributes);

        [Description("Sets the date and time the file was created.")]
        void SetCreationTime(string path, DateTime creationTime);

        [Description("Sets the date and time, in Coordinated Universal Time (UTC), that the file was created.")]
        void SetCreationTimeUtc(string path, DateTime creationTimeUtc);

        [Description("Sets the date and time the specified file was last accessed.")]
        void SetLastAccessTime(string path, DateTime lastAccessTime);

        [Description("Sets the date and time, in Coordinated Universal Time (UTC), that the specified file was last accessed.")]
        void SetLastAccessTimeUtc(string path, DateTime lastAccessTimeUtc);

        [Description("Sets the date and time that the specified file was last written to.")]
        void SetLastWriteTime(string path, DateTime lastWriteTime);

        [Description("Sets the date and time, in Coordinated Universal Time (UTC), that the specified file was last written to.")]
        void SetLastWriteTimeUtc(string path, DateTime lastWriteTimeUtc);

        [Description("Creates a new file, writes the specified byte array to the file, and then closes the file. If the target file already exists, it is overwritten.")]
        void WriteAllBytes(string path, [In] ref byte[] bytes);

        [Description("Creates a new file, write the specified string array to the file, and then closes the file.")]
        void WriteAllLines(string path, [In] ref string[] contents);

        [Description("Creates a new file, writes the specified string array to the file by using the specified encoding, and then closes the file.")]
        void WriteAllLines(string path, [In] ref string[] contents, Encoding encoding);

        [Description("Creates a new file, writes a collection of strings to the file, and then closes the file.")]
        void WriteAllLines(string path, GCollections.IEnumerable contents);

        [Description("Creates a new file by using the specified encoding, writes a collection of strings to the file, and then closes the file.")]
        void WriteAllLines(string path, GCollections.IEnumerable contents, Encoding encoding);

        [Description("Creates a new file, writes the specified string to the file, and then closes the file. If the target file already exists, it is overwritten.")]
        void WriteAllText(string path, string contents);

        [Description("Creates a new file, writes the specified string to the file using the specified encoding, and then closes the file. If the target file already exists, it is overwritten.")]
        void WriteAllText(string path, string contents, Encoding encoding);
    }
}
