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
    [Description("Exposes static methods for creating, moving, and enumerating through directories and subdirectories. This class cannot be inherited.")]
    [Guid("377FB369-8DC9-4848-8D70-F5FE6F24C446")]
    [ProgId("DotNetLib.System.IO.DirectorySingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDirectorySingleton))]

    public class DirectorySingleton : IDirectorySingleton
    {
        public DirectorySingleton() { }

        public DirectoryInfo CreateDirectory(string path)
        {
            return new DirectoryInfo(GIO.Directory.CreateDirectory(path));
        }

        public DirectoryInfo CreateDirectory(string path, GAccessControl.DirectorySecurity directorySecurity)
        {
            return new DirectoryInfo(GIO.Directory.CreateDirectory(path, directorySecurity));
        }

        public void Delete(string path)
        {
            GIO.Directory.Delete(path);
        }

        public void Delete(string path, bool recursive)
        {
            GIO.Directory.Delete(path, recursive);
        }

        public GCollections.IEnumerable EnumerateDirectories(string path, string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly)
        {
            return GIO.Directory.EnumerateDirectories(path, searchPattern, searchOption);
        }

        public GCollections.IEnumerable EnumerateFiles(string path, string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly)
        {
            return GIO.Directory.EnumerateFiles(path, searchPattern, searchOption);
        }

        public GCollections.IEnumerable EnumerateFileSystemEntries(string path, string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly)
        {
            return GIO.Directory.EnumerateFileSystemEntries(path, searchPattern, searchOption);
        }

        public bool Exists(string path)
        {
            return GIO.Directory.Exists(path);
        }

        public GAccessControl.DirectorySecurity GetAccessControl(string path)
        {
            return GIO.Directory.GetAccessControl(path);
        }

        public GAccessControl.DirectorySecurity GetAccessControl(string path, AccessControlSections includeSections)
        {
            return GIO.Directory.GetAccessControl(path, (GAccessControl.AccessControlSections)includeSections);
        }

        public DateTime GetCreationTime(string path)
        {
            return new DateTime(GIO.Directory.GetCreationTime(path));
        }

        public DateTime GetCreationTimeUtc(string path)
        {
            return new DateTime(GIO.Directory.GetCreationTimeUtc(path));
        }

        public string GetCurrentDirectory()
        {
            return GIO.Directory.GetCurrentDirectory();
        }

        public string[] GetDirectories(string path, string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly)
        {
            return GIO.Directory.GetDirectories(path, searchPattern, searchOption);
        }

        public string GetDirectoryRoot(string path)
        {
            return GIO.Directory.GetDirectoryRoot(path);
        }

        public string[] GetFiles(string path, string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly)
        {
            return GIO.Directory.GetFiles(path, searchPattern, searchOption);
        }

        public string[] GetFileSystemEntries(string path, string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly)
        {
            return GIO.Directory.GetFileSystemEntries(path, searchPattern, searchOption);
        }

        public DateTime GetLastAccessTime(string path)
        {
            return new DateTime(GIO.Directory.GetLastAccessTime(path));
        }

        public DateTime GetLastAccessTimeUtc(string path)
        {
            return new DateTime(GIO.Directory.GetLastAccessTimeUtc(path));
        }

        public DateTime GetLastWriteTime(string path)
        {
            return new DateTime(GIO.Directory.GetLastWriteTime(path));
        }

        public DateTime GetLastWriteTimeUtc(string path)
        {
            return new DateTime(GIO.Directory.GetLastWriteTimeUtc(path));
        }

        public string[] GetLogicalDrives()
        {
            return GIO.Directory.GetLogicalDrives();
        }

        public DirectoryInfo GetParent(string path)
        {
            return new DirectoryInfo(GIO.Directory.GetParent(path));
        }

        public void Move(string sourceDirName, string destDirName)
        {
            GIO.Directory.Move(sourceDirName, destDirName);
        }

        public void SetAccessControl(string path, GAccessControl.DirectorySecurity directorySecurity)
        {
            GIO.Directory.SetAccessControl(path, directorySecurity);
        }

        public void SetCreationTime(string path, DateTime creationTime)
        {
            GIO.Directory.SetCreationTime(path, creationTime.WrappedDateTime);
        }

        public void SetCreationTimeUtc(string path, DateTime creationTimeUtc)
        {
            GIO.Directory.SetCreationTimeUtc(path, creationTimeUtc.WrappedDateTime);
        }

        public void SetCurrentDirectory(string path)
        {
            GIO.Directory.SetCurrentDirectory(path);
        }

        public void SetLastAccessTime(string path, DateTime lastAccessTime)
        {
            GIO.Directory.SetLastAccessTime(path, lastAccessTime.WrappedDateTime);
        }

        public void SetLastAccessTimeUtc(string path, DateTime lastAccessTimeUtc)
        {
            GIO.Directory.SetLastAccessTimeUtc(path, lastAccessTimeUtc.WrappedDateTime);
        }

        public void SetLastWriteTime(string path, DateTime lastWriteTime)
        {
            GIO.Directory.SetLastWriteTime(path, lastWriteTime.WrappedDateTime);
        }

        public void SetLastWriteTimeUtc(string path, DateTime lastWriteTimeUtc)
        {
            GIO.Directory.SetLastWriteTimeUtc(path, lastWriteTimeUtc.WrappedDateTime);
        }

    }
}
