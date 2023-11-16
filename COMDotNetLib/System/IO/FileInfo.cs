// https://learn.microsoft.com/en-us/dotnet/api/system.io.fileinfo?view=netframework-4.8.1

using GSystem = global::System;
using GIO = global::System.IO;
using GAccessControl = global::System.Security.AccessControl;
using GSerialization = global::System.Runtime.Serialization;

using DotNetLib.System.Security.AccessControl;
using System.IO;
using DotNetLib.Extensions;
using System.Runtime.Serialization;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("")]
    [Guid("07C5794D-221C-4ED1-8DAF-AA7249E2CFFE")]
    [ProgId("DotNetLib.System.IO.FileInfo")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IFileInfo))]
    public class FileInfo : IFileInfo, FileSystemInfo, ISerializable,  IWrappedObject 
    {
        private GIO.FileInfo _fileInfo;

        public FileInfo(string fileName)
        {
            _fileInfo = new GIO.FileInfo(fileName);
        }

        public FileInfo(GIO.FileInfo fileInfo)
        {
            _fileInfo = fileInfo;
        }

        // Properties

        public GIO.FileInfo WrappedDirectoryInfo
        {
            get => _fileInfo;
            set => _fileInfo = value; 
        }

        public object WrappedObject => _fileInfo;

        public DirectoryInfo Directory => new DirectoryInfo(_fileInfo.Directory);

        public string DirectoryName => _fileInfo.DirectoryName;

        public bool Exists => _fileInfo.Exists;

        public bool IsReadOnly 
        {
            get => _fileInfo.IsReadOnly; 
            set =>  _fileInfo.IsReadOnly = value;
        }

        public long Length => _fileInfo.Length;

        public string Name => _fileInfo.Name;

        // Inherited from FileSystemInfo
        public GIO.FileAttributes Attributes 
        { 
            get => _fileInfo.Attributes;
            set => _fileInfo.Attributes = value;
        }

        public DateTime CreationTime 
        {
            get => new DateTime(_fileInfo.CreationTime);
            set => _fileInfo.CreationTime = value.WrappedDateTime;
        }

        public DateTime CreationTimeUtc 
        {
            get => new DateTime(_fileInfo.CreationTimeUtc);
            set => _fileInfo.CreationTimeUtc = value.WrappedDateTime;
        }

        public string Extension => _fileInfo.Extension;

        public string FullName => _fileInfo.FullName;

        public DateTime LastAccessTime 
        { 
            get => new DateTime(_fileInfo.LastAccessTime);
            set => _fileInfo.LastAccessTime = value.WrappedDateTime;
        }

        public DateTime LastAccessTimeUtc 
        {
            get => new DateTime(_fileInfo.LastAccessTimeUtc);
            set => _fileInfo.LastAccessTimeUtc = value.WrappedDateTime;
        }

        public DateTime LastWriteTime 
        {
            get => new DateTime(_fileInfo.LastWriteTime);
            set => _fileInfo.LastWriteTime = value.WrappedDateTime;
        }

        public DateTime LastWriteTimeUtc 
        {
            get => new DateTime(_fileInfo.LastWriteTimeUtc);
            set => _fileInfo.LastWriteTimeUtc = value.WrappedDateTime;
        }


        //Methods

        public GIO.StreamWriter AppendText()
        {
            return _fileInfo.AppendText();
        }

        public FileInfo CopyTo(string destFileName)
        {
            return new FileInfo(_fileInfo.CopyTo(destFileName));
        }

        public FileInfo CopyTo(string destFileName, bool overwrite)
        {
            return new FileInfo(_fileInfo.CopyTo(destFileName,overwrite));
        }

        public GIO.FileStream Create()
        {
            return _fileInfo.Create();
        }

        //public virtual System.Runtime.Remoting.ObjRef CreateObjRef (Type requestedType);


        public GIO.StreamWriter CreateText()
        {
            return _fileInfo.CreateText();
        }

        public void Decrypt()
        { 
            _fileInfo.Decrypt();
        }

        public void Delete()
        { 
            _fileInfo.Delete(); 
        }

        public void Encrypt()
        {
            _fileInfo.Encrypt();
        }

        new public virtual bool Equals(object obj)
        {
            return Equals(obj.Unwrap());
        }

        public GAccessControl.FileSecurity GetAccessControl()
        {
            return _fileInfo.GetAccessControl();
        }

        public GAccessControl.FileSecurity GetAccessControl(AccessControlSections includeSections)
        {
            return _fileInfo.GetAccessControl((GAccessControl.AccessControlSections)includeSections);
        }

        new public virtual int GetHashCode()
        { 
            return _fileInfo.GetHashCode(); 
        }

        // public object GetLifetimeService ();

        public virtual void GetObjectData(GSerialization.SerializationInfo info, GSerialization.StreamingContext context)
        {
            _fileInfo.GetObjectData(info, context);
        }

        public new Type GetType()
        {
            return new Type(((GSystem.Object)this).GetType());
        }

        public void MoveTo(string destFileName)
        {
            _fileInfo.MoveTo(destFileName);
        }

        public GIO.FileStream Open(GIO.FileMode mode)
        { 
            return _fileInfo.Open(mode); 
        }

        public GIO.FileStream Open(GIO.FileMode mode, GIO.FileAccess access)
        {
            return _fileInfo.Open(mode, access);
        }

        public GIO.FileStream Open(GIO.FileMode mode, GIO.FileAccess access, GIO.FileShare share)
        {
            return _fileInfo.Open(mode, access, share);
        }

        public GIO.FileStream OpenRead()
        { 
            return _fileInfo.OpenRead(); 
        }

        public GIO.StreamReader OpenText()
        { 
            return _fileInfo.OpenText();
        }

        public GIO.FileStream OpenWrite()
        { 
            return _fileInfo.OpenWrite(); 
        }

        public void Refresh()
        { 
            _fileInfo.Refresh(); 
        }

        public FileInfo Replace(string destinationFileName, string destinationBackupFileName)
        {
            return  new FileInfo(_fileInfo.Replace(destinationFileName, destinationBackupFileName));
        }

        public FileInfo Replace(string destinationFileName, string destinationBackupFileName, bool ignoreMetadataErrors)
        {
            return new FileInfo(_fileInfo.Replace(destinationFileName,destinationBackupFileName,ignoreMetadataErrors));
        }

        public void SetAccessControl(GAccessControl.FileSecurity fileSecurity)
        {
            _fileInfo.SetAccessControl(fileSecurity);
        }

        public override string ToString() 
        { 
            return _fileInfo.ToString();
        }

    }
}
