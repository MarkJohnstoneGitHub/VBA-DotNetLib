// https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo?view=netframework-4.8.1

using GCollections = global::System.Collections;
using GIO = global::System.IO;
using GAccessControl = global::System.Security.AccessControl;
using System.ComponentModel;
using DotNetLib.System.Security.AccessControl;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("5BCD9DAD-4AD9-4C03-9D36-4FD7306D2D69")]
    [Description("Exposes instance methods for creating, moving, and enumerating through directories and subdirectories. This class cannot be inherited.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDirectoryInfo
    {
        bool Exists 
        {
            [Description("Gets a value indicating whether the directory exists.")]
            get;
        }

        string FullName 
        {
            [Description("Gets the full path of the directory.")]
            get;
        }

        string Name 
        {
            [Description("Gets the name of this DirectoryInfo instance.")]
            get;
        }


        DirectoryInfo Parent 
        {
            [Description("Gets the parent directory of a specified subdirectory.")]
            get;
        }

        DirectoryInfo Root 
        {
            [Description("Gets the root portion of the directory")]
            get;
        }

        //  Methods
        [Description("Creates a directory.")]
        void Create();

        [Description("Creates a directory using a DirectorySecurity object.")]
        void Create(GAccessControl.DirectorySecurity directorySecurity);

        [Description("Creates a subdirectory or subdirectories on the specified path. The specified path can be relative to this instance of the DirectoryInfo class.")]
        DirectoryInfo CreateSubdirectory(string path);

        [Description("Creates a subdirectory or subdirectories on the specified path with the specified security. The specified path can be relative to this instance of the DirectoryInfo class.")]
        DirectoryInfo CreateSubdirectory(string path, GAccessControl.DirectorySecurity directorySecurity = null);

        //[Description("Deletes this DirectoryInfo if it is empty.")]
        //void Delete();

        [Description("Deletes this instance of a DirectoryInfo, specifying whether to delete subdirectories and files.")]
        void Delete(bool recursive = false);

        [Description("Returns an enumerable collection of directory information that matches a specified search pattern and search subdirectory option.")]
        GCollections.IEnumerable EnumerateDirectories(string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly);

        [Description("Returns an enumerable collection of file information that matches a specified search pattern and search subdirectory option.")]
        GCollections.IEnumerable EnumerateFiles(string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly);

        [Description("Returns an enumerable collection of file system information that matches a specified search pattern and search subdirectory option.")]
        GCollections.IEnumerable EnumerateFileSystemInfos(string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly);

        [Description("Gets a DirectorySecurity object that encapsulates the access control list (ACL) entries for the directory described by the current DirectoryInfo object.")]
        GAccessControl.DirectorySecurity GetAccessControl();

        [Description("Gets a DirectorySecurity object that encapsulates the specified type of access control list (ACL) entries for the directory described by the current DirectoryInfo object.")]
        GAccessControl.DirectorySecurity GetAccessControl(AccessControlSections includeSections);

        [Description("Returns an array of directories in the current DirectoryInfo matching the given search criteria and using a value to determine whether to search subdirectories.")]
        DirectoryInfo[] GetDirectories(string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly);

        [Description("Returns a file list from the current directory matching the given search pattern and using a value to determine whether to search subdirectories.")]
        FileInfo[] GetFiles(string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly);

        [Description("Retrieves an array of FileSystemInfo objects that represent the files and subdirectories matching the specified search criteria.")]
        FileSystemInfo[] GetFileSystemInfos(string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly);

        [Description("Moves a DirectoryInfo instance and its contents to a new path.")]
        void MoveTo(string destDirName);

        [Description("Applies access control list (ACL) entries described by a DirectorySecurity object to the directory described by the current DirectoryInfo object.")]
        void SetAccessControl(GAccessControl.DirectorySecurity directorySecurity);

        [Description("Returns the original path that was passed to the DirectoryInfo constructor. Use the FullName or Name properties for the full path or file/directory name instead of this method.")]
        string ToString();


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

        //string FullName
        //{
        //    [Description("Gets the full path of the directory or file.")]
        //    get;
        //}

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


    }
}
