// https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo?view=netframework-4.8.1

using GCollections = global::System.Collections;
using GIO = global::System.IO;
using GAccessControl = global::System.Security.AccessControl;
using DotNetLib.System.Security.AccessControl;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using DotNetLib.Extensions;
using System.ComponentModel;
using System.Runtime.Serialization;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("Exposes instance methods for creating, moving, and enumerating through directories and subdirectories. This class cannot be inherited.")]
    [Guid("C14198C4-6A2C-4089-BC05-04911769E735")]
    [ProgId("DotNetLib.System.IO.DirectoryInfo")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDirectoryInfo))]
    public class DirectoryInfo :  IDirectoryInfo, FileSystemInfo
    {
        private GIO.DirectoryInfo _directoryInfo;

        public DirectoryInfo(string path)
        {
            _directoryInfo = new GIO.DirectoryInfo(path);
        }

        public DirectoryInfo(GIO.DirectoryInfo directoryInfo)
        {
            _directoryInfo = directoryInfo;
        }

        //Properties

        public GIO.DirectoryInfo WrappedDirectoryInfo
        {
            get { return _directoryInfo; }
            set { _directoryInfo = value; }
        }

        public bool Exists => _directoryInfo.Exists;

        public string FullName => _directoryInfo.FullName;

        public string Name => _directoryInfo.Name;

        public DirectoryInfo Parent => new DirectoryInfo(_directoryInfo.Parent);

        public DirectoryInfo Root => new DirectoryInfo(_directoryInfo.Root);

        public FileAttributes Attributes 
        {
            get => _directoryInfo.Attributes;
            set => _directoryInfo.Attributes = value;
        }

        public DateTime CreationTime 
        {
            get => new DateTime(_directoryInfo.CreationTime);
            set => _directoryInfo.CreationTime = value.WrappedDateTime;
        }

        public DateTime CreationTimeUtc
        {
            get => new DateTime(_directoryInfo.CreationTimeUtc);
            set => _directoryInfo.CreationTimeUtc = value.WrappedDateTime;
        }

        public string Extension => _directoryInfo.Extension;

        public DateTime LastAccessTime
        {
            get => new DateTime(_directoryInfo.LastAccessTime);
            set => _directoryInfo.LastAccessTime = value.WrappedDateTime;
        }

        public DateTime LastAccessTimeUtc
        {
            get => new DateTime(_directoryInfo.LastAccessTimeUtc);
            set => _directoryInfo.LastAccessTimeUtc = value.WrappedDateTime;
        }

        public DateTime LastWriteTime
        {
            get => new DateTime(_directoryInfo.LastWriteTime);
            set => _directoryInfo.LastWriteTime = value.WrappedDateTime;
        }

        public DateTime LastWriteTimeUtc
        {
            get => new DateTime(_directoryInfo.LastWriteTimeUtc);
            set => _directoryInfo.LastWriteTimeUtc = value.WrappedDateTime;
        }

        // Methods
        public void Create()
        {
            _directoryInfo.Create();
        }

        public void Create(GAccessControl.DirectorySecurity directorySecurity)
        {
            _directoryInfo.Create(directorySecurity);
        }

        public DirectoryInfo CreateSubdirectory(string path)
        {
            return new DirectoryInfo(_directoryInfo.CreateSubdirectory(path));
        }

        public DirectoryInfo CreateSubdirectory(string path, GAccessControl.DirectorySecurity directorySecurity = null)
        {
            return new DirectoryInfo(_directoryInfo.CreateSubdirectory(path,directorySecurity));
        }

        public void Delete()
        { 
            _directoryInfo.Delete(); 
        }

        public void Delete(bool recursive)
        { 
            _directoryInfo.Delete(recursive); 
        }

        public GCollections.IEnumerable EnumerateDirectories(string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly)
        {
            IEnumerable<GIO.DirectoryInfo> directoryInfoList = _directoryInfo.EnumerateDirectories(searchPattern, searchOption);

            List <DirectoryInfo> wrappedDirectoryInfoList = new List<DirectoryInfo>(directoryInfoList.Count());
            foreach (GIO.DirectoryInfo directoryInfo in directoryInfoList)
            { 
                wrappedDirectoryInfoList.Add(new DirectoryInfo(directoryInfo));
            }
            return wrappedDirectoryInfoList;
        }

        public GCollections.IEnumerable EnumerateFiles(string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly)
        {
            IEnumerable<GIO.FileInfo> fileInfoList = _directoryInfo.EnumerateFiles(searchPattern, searchOption);

            List<FileInfo> wrappedFileInfoList = new List<FileInfo>(fileInfoList.Count());
            foreach (GIO.FileInfo fileInfo in fileInfoList)
            {
                wrappedFileInfoList.Add(new FileInfo(fileInfo));
            }
            return wrappedFileInfoList;
        }

        public GCollections.IEnumerable EnumerateFileSystemInfos(string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly)
        {
            IEnumerable<GIO.FileSystemInfo> fileInfoList = _directoryInfo.EnumerateFileSystemInfos(searchPattern, searchOption);
            List<FileSystemInfo> wrappedFileSystemInfoList = new List<FileSystemInfo>(fileInfoList.Count());
            foreach (GIO.FileSystemInfo fileSystemInfo in fileInfoList)
            {
                wrappedFileSystemInfoList.Add(fileSystemInfo.Wrap());
            }
            return wrappedFileSystemInfoList;
        }

        public GAccessControl.DirectorySecurity GetAccessControl()
        { 
            return _directoryInfo.GetAccessControl(); 
        }

        public GAccessControl.DirectorySecurity GetAccessControl(AccessControlSections includeSections)
        { 
            return _directoryInfo.GetAccessControl((GAccessControl.AccessControlSections)includeSections); 
        }

        public DirectoryInfo[] GetDirectories(string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly)
        {
            return WrapDirectoryInfoArray(_directoryInfo.GetDirectories());
        }

        // Todo
        public FileInfo[] GetFiles(string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly)
        {
            return WrapFileInfofoArray(_directoryInfo.GetFiles(searchPattern, searchOption));
        }

        // Todo
        public FileSystemInfo[] GetFileSystemInfos(string searchPattern = "*", GIO.SearchOption searchOption = GIO.SearchOption.TopDirectoryOnly)
        { 
            return WrapFileSystemInfoArray(_directoryInfo.GetFileSystemInfos(searchPattern, searchOption)); 
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            _directoryInfo.GetObjectData(info, context);
        }

        public void MoveTo(string destDirName)
        { 
            _directoryInfo.MoveTo(destDirName);
        }

        public void Refresh()
        {
            _directoryInfo.Refresh();
        }

        public void SetAccessControl(GAccessControl.DirectorySecurity directorySecurity)
        {
            _directoryInfo.SetAccessControl(directorySecurity);
        }

        public override string ToString()
        { 
            return _directoryInfo.ToString(); 
        }

        public DirectoryInfo[] WrapDirectoryInfoArray(GIO.DirectoryInfo[] directoryInfoArray)
        {
            if (directoryInfoArray == null)
                return null;

            DirectoryInfo[] wrappedDirectoryInfoArray = new DirectoryInfo[directoryInfoArray.Length];
            for (int index = 0; index < wrappedDirectoryInfoArray.Length; index++)
            {
                wrappedDirectoryInfoArray[index] = new DirectoryInfo(directoryInfoArray[index]);
            }
            return wrappedDirectoryInfoArray;
        }

        public FileInfo[] WrapFileInfofoArray(GIO.FileInfo[] fileInfoArray)
        {
            if (fileInfoArray == null)
                return null;

            FileInfo[] wrappedFileInfoArray = new FileInfo[fileInfoArray.Length];
            for (int index = 0; index < wrappedFileInfoArray.Length; index++)
            {
                wrappedFileInfoArray[index] = new FileInfo(fileInfoArray[index]);
            }
            return wrappedFileInfoArray;
        }

        public FileSystemInfo[] WrapFileSystemInfoArray(GIO.FileSystemInfo[] fileSystemInfoArray)
        {
            if (fileSystemInfoArray == null)
                return null;

            FileSystemInfo[] wrappedFileSystemInfoArray = new FileSystemInfo[fileSystemInfoArray.Length];
            for (int index = 0; index < wrappedFileSystemInfoArray.Length; index++)
            {
                wrappedFileSystemInfoArray[index] = fileSystemInfoArray[index].Wrap();
            }
            return wrappedFileSystemInfoArray;
        }


        //public GCollections.IEnumerable EnumerateDirectories()
        //{ 
        //    return _directoryInfo.EnumerateDirectories(); 
        //}

        //public GCollections.IEnumerable EnumerateDirectories(string searchPattern)
        //{ 
        //    return _directoryInfo.EnumerateDirectories(searchPattern); 
        //}

        //public GCollections.IEnumerable EnumerateFiles()
        //{ 
        //    return _directoryInfo.EnumerateFiles(); 
        //}

        //public GCollections.IEnumerable EnumerateFiles(string searchPattern)
        //{ 
        //    return _directoryInfo.EnumerateFiles(searchPattern); 
        //}


        //public GCollections.IEnumerable EnumerateFileSystemInfos()
        //{ 
        //    return _directoryInfo.EnumerateFileSystemInfos();
        //}

        //public GCollections.IEnumerable EnumerateFileSystemInfos(string searchPattern)
        //{
        //    return _directoryInfo.EnumerateFileSystemInfos(searchPattern);
        //}


        //public GIO.DirectoryInfo[] GetDirectories()
        //{ 
        //    return _directoryInfo.GetDirectories(); 
        //}

        //public GIO.DirectoryInfo[] GetDirectories(string searchPattern)
        //{
        //    return _directoryInfo.GetDirectories(searchPattern);
        //}

        //public GIO.DirectoryInfo[] GetDirectories(string searchPattern, GIO.SearchOption searchOption)
        //{ 
        //    return _directoryInfo.GetDirectories(searchPattern, searchOption); 
        //}



    }
}
