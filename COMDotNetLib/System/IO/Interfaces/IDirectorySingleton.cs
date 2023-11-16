
// https://learn.microsoft.com/en-us/dotnet/api/system.io.directory?view=netframework-4.8.1

using GCollections = global::System.Collections;
using GIO = global::System.IO;
using GAccessControl = global::System.Security.AccessControl;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.System.Security.AccessControl;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("4A061E3C-6F23-45C7-90C2-0F1BE089F2C3")]
    [Description("Exposes static methods for creating, moving, and enumerating through directories and subdirectories. This class cannot be inherited.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDirectorySingleton
    {
        [Description("Creates all directories and subdirectories in the specified path unless they already exist.")]
        DirectoryInfo CreateDirectory(string path);

        [Description("Creates all the directories in the specified path, unless they already exist, applying the specified Windows security.")]
        DirectoryInfo CreateDirectory(string path, GAccessControl.DirectorySecurity directorySecurity);

        [Description("Deletes an empty directory from a specified path.")]
        void Delete(string path);

        [Description("Deletes the specified directory and, if indicated, any subdirectories and files in the directory.")]
        void Delete(string path, bool recursive);

        [Description("Returns an enumerable collection of directory full names that match a search pattern in a specified path, and optionally searches subdirectories.")]
        GCollections.IEnumerable EnumerateDirectories(string path, string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly);

        [Description("Returns an enumerable collection of full file names that match a search pattern in a specified path, and optionally searches subdirectories.")]
        GCollections.IEnumerable EnumerateFiles(string path, string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly);

        [Description("Returns an enumerable collection of file names and directory names that match a search pattern in a specified path, and optionally searches subdirectories.")]
        GCollections.IEnumerable EnumerateFileSystemEntries(string path, string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly);

        [Description("Determines whether the given path refers to an existing directory on disk.")]
        bool Exists(string path);

        [Description("Gets a DirectorySecurity object that encapsulates the access control list (ACL) entries for a specified directory")]
        GAccessControl.DirectorySecurity GetAccessControl(string path);

        [Description("Gets a DirectorySecurity object that encapsulates the specified type of access control list (ACL) entries for a specified directory.")]
        GAccessControl.DirectorySecurity GetAccessControl(string path, AccessControlSections includeSections);

        [Description("Gets the creation date and time of a directory.")]
        DateTime GetCreationTime(string path);

        [Description("Gets the creation date and time, in Coordinated Universal Time (UTC) format, of a directory.")]
        DateTime GetCreationTimeUtc(string path);

        [Description("Gets the current working directory of the application.")]
        string GetCurrentDirectory();

        [Description("Returns the names of the subdirectories (including their paths) that match the specified search pattern in the specified directory, and optionally searches subdirectories.")]
        string[] GetDirectories(string path, string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly);

        [Description("Returns the volume information, root information, or both for the specified path.")]
        string GetDirectoryRoot(string path);

        [Description("Returns the names of files (including their paths) that match the specified search pattern in the specified directory, using a value to determine whether to search subdirectories.")]
        string[] GetFiles(string path, string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly);

        [Description("Returns an array of all the file names and directory names that match a search pattern in a specified path, and optionally searches subdirectories.")]
        string[] GetFileSystemEntries(string path, string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly);

        [Description("Returns the date and time the specified file or directory was last accessed.")]
        DateTime GetLastAccessTime(string path);

        [Description("Returns the date and time, in Coordinated Universal Time (UTC) format, that the specified file or directory was last accessed.")]
        DateTime GetLastAccessTimeUtc(string path);

        [Description("Returns the date and time the specified file or directory was last written to.")]
        DateTime GetLastWriteTime(string path);

        [Description("Returns the date and time, in Coordinated Universal Time (UTC) format, that the specified file or directory was last written to.")]
        DateTime GetLastWriteTimeUtc(string path);

        [Description("Retrieves the names of the logical drives on this computer in the form \"<drive letter>:\\\".")]
        string[] GetLogicalDrives();

        [Description("Retrieves the parent directory of the specified path, including both absolute and relative paths.")]
        DirectoryInfo GetParent(string path);

        [Description("Moves a file or a directory and its contents to a new location.")]
        void Move(string sourceDirName, string destDirName);

        [Description("Applies access control list (ACL) entries described by a DirectorySecurity object to the specified directory.")]
        void SetAccessControl(string path, GAccessControl.DirectorySecurity directorySecurity);

        [Description("Sets the creation date and time for the specified file or directory.")]
        void SetCreationTime(string path, DateTime creationTime);

        [Description("Sets the creation date and time, in Coordinated Universal Time (UTC) format, for the specified file or directory.")]
        void SetCreationTimeUtc(string path, DateTime creationTimeUtc);

        [Description("Sets the application's current working directory to the specified directory.")]
        void SetCurrentDirectory(string path);

        [Description("Sets the date and time the specified file or directory was last accessed.")]
        void SetLastAccessTime(string path, DateTime lastAccessTime);

        [Description("Sets the date and time, in Coordinated Universal Time (UTC) format, that the specified file or directory was last accessed.")]
        void SetLastAccessTimeUtc(string path, DateTime lastAccessTimeUtc);

        [Description("Sets the date and time a directory was last written to.")]
        void SetLastWriteTime(string path, DateTime lastWriteTime);

        [Description("Sets the date and time, in Coordinated Universal Time (UTC) format, that a directory was last written to.")]
        void SetLastWriteTimeUtc(string path, DateTime lastWriteTimeUtc);
    }
}
