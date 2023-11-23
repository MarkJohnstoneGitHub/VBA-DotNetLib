// https://learn.microsoft.com/en-us/dotnet/api/system.io.file?view=netframework-4.8.1

using Encoding = DotNetLib.System.Text.Encoding;
using GText = global::System.Text;
using GCollections = global::System.Collections;
using GIO = global::System.IO;
using GAccessControl = global::System.Security.AccessControl;


using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.System.Security.AccessControl;
using System.Text;
using DotNetLib.Extensions;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("Provides static methods for the creation, copying, deletion, moving, and opening of a single file, and aids in the creation of FileStream objects.")]
    [Guid("E6EB8B4D-D52C-44AD-9D48-08C766D9E8B1")]
    [ProgId("DotNetLib.System.IO.FileSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IFileSingleton))]
    public class FileSingleton : IFileSingleton
    {
        public FileSingleton() { }

        public void AppendAllLines(string path, GCollections.IEnumerable contents)
        {
            GIO.File.AppendAllLines(path, (GCollections.Generic.IEnumerable<string>)contents);
        }

        public void AppendAllLines(string path, GCollections.IEnumerable contents, Encoding encoding)
        {
            GIO.File.AppendAllLines(path, (GCollections.Generic.IEnumerable<string>)contents, encoding.UnWrapEncoding());
        }

        public void AppendAllText(string path, string contents)
        {
            GIO.File.AppendAllText(path, contents);
        }

        public void AppendAllText(string path, string contents, Encoding encoding)
        {
            GIO.File.AppendAllText(path, contents, encoding.UnWrapEncoding());
        }

        public StreamWriter AppendText(string path)
        {
           return new StreamWriter(GIO.File.AppendText(path));
        }

        public void Copy(string sourceFileName, string destFileName)
        {
            GIO.File.Copy(sourceFileName, destFileName);
        }

        public void Copy(string sourceFileName, string destFileName, bool overwrite)
        {
            GIO.File.Copy(sourceFileName, destFileName, overwrite);
        }

        public GIO.FileStream Create(string path)
        {  
            return GIO.File.Create(path);
        }

        public GIO.FileStream Create(string path, int bufferSize)
        {
            return GIO.File.Create(path, bufferSize);
        }

        public GIO.FileStream Create(string path, int bufferSize, GIO.FileOptions options)
        {
            return GIO.File.Create(path, bufferSize, options);
        }

        public GIO.FileStream Create(string path, int bufferSize, GIO.FileOptions options, GAccessControl.FileSecurity fileSecurity)
        {
            return GIO.File.Create(path, bufferSize, options, fileSecurity);
        }

        public StreamWriter CreateText(string path)
        {
            return new StreamWriter(GIO.File.CreateText(path));
        }

        public void Decrypt(string path)
        { 
            GIO.File.Decrypt(path);
        }

        public void Delete(string path) 
        { 
            GIO.File.Delete(path);
        }

        public void Encrypt(string path)
        { 
            GIO.File.Encrypt(path);
        }

        public bool Exists(string path)
        {
            return GIO.File.Exists(path);
        }

        public GAccessControl.FileSecurity GetAccessControl(string path)
        {
            return GIO.File.GetAccessControl(path);
        }

        public GAccessControl.FileSecurity GetAccessControl(string path, AccessControlSections includeSections)
        {
            return GIO.File.GetAccessControl(path, (GAccessControl.AccessControlSections)includeSections);
        }

        public GIO.FileAttributes GetAttributes(string path)
        {
            return GIO.File.GetAttributes(path);
        }

        public DateTime GetCreationTime(string path)
        {
            return new DateTime(GIO.File.GetCreationTime(path));
        }

        public DateTime GetCreationTimeUtc(string path)
        {
            return new DateTime(GIO.File.GetCreationTimeUtc(path));
        }

        public DateTime GetLastAccessTime(string path)
        {
            return new DateTime(GIO.File.GetLastAccessTime(path));
        }

        public DateTime GetLastAccessTimeUtc(string path)
        {
            return new DateTime(GIO.File.GetLastAccessTimeUtc(path));
        }

        public DateTime GetLastWriteTime(string path)
        {
            return new DateTime(GIO.File.GetLastWriteTime(path));
        }

        public DateTime GetLastWriteTimeUtc(string path)
        {
            return new DateTime(GIO.File.GetLastWriteTimeUtc(path));
        }

        public void Move(string sourceFileName, string destFileName)
        {
            GIO.File.Move(sourceFileName, destFileName);
        }

        public GIO.FileStream Open(string path, GIO.FileMode mode)
        {
            return GIO.File.Open(path, mode);
        }

        public GIO.FileStream Open(string path, GIO.FileMode mode, GIO.FileAccess access)
        {
            return GIO.File.Open(path, mode, access);
        }

        public GIO.FileStream Open(string path, GIO.FileMode mode, GIO.FileAccess access, GIO.FileShare share)
        {
            return GIO.File.Open(path, mode, access, share);
        }

        public GIO.FileStream OpenRead(string path)
        {
            return GIO.File.OpenRead(path);
        }

        public GIO.StreamReader OpenText(string path)
        {
            return GIO.File.OpenText(path);
        }

        public GIO.FileStream OpenWrite(string path)
        {
            return GIO.File.OpenWrite(path);
        }

        public byte[] ReadAllBytes(string path)
        {
            return GIO.File.ReadAllBytes(path);
        }

        public string[] ReadAllLines(string path)
        {
            return GIO.File.ReadAllLines(path);
        }

        public string[] ReadAllLines(string path, Encoding encoding)
        {
            return GIO.File.ReadAllLines(path, encoding.UnWrapEncoding());
        }

        public string ReadAllText(string path)
        {
            return GIO.File.ReadAllText(path);
        }

        public string ReadAllText(string path, Encoding encoding)
        {
            return GIO.File.ReadAllText(path, encoding.UnWrapEncoding());
        }

        public GCollections.IEnumerable ReadLines(string path)
        {
            return GIO.File.ReadLines(path);
        }

        public GCollections.IEnumerable ReadLines(string path, Encoding encoding)
        {
            return GIO.File.ReadLines(path, encoding.UnWrapEncoding());
        }

        public void Replace(string sourceFileName, string destinationFileName, string destinationBackupFileName)
        {
            GIO.File.Replace(sourceFileName, destinationFileName, destinationBackupFileName);
        }

        public void Replace(string sourceFileName, string destinationFileName, string destinationBackupFileName, bool ignoreMetadataErrors)
        {
            GIO.File.Replace(sourceFileName, destinationFileName, destinationBackupFileName, ignoreMetadataErrors);
        }

        public void SetAccessControl(string path, GAccessControl.FileSecurity fileSecurity)
        {
            GIO.File.SetAccessControl(path, fileSecurity);
        }

        public void SetAttributes(string path, GIO.FileAttributes fileAttributes)
        {
            GIO.File.SetAttributes(path, fileAttributes);
        }

        public void SetCreationTime(string path, DateTime creationTime)
        {
            GIO.File.SetCreationTime(path, creationTime.WrappedDateTime);
        }

        public void SetCreationTimeUtc(string path, DateTime creationTimeUtc)
        {
            GIO.File.SetCreationTimeUtc(path, creationTimeUtc.WrappedDateTime);
        }

        public void SetLastAccessTime(string path, DateTime lastAccessTime)
        {
            GIO.File.SetLastAccessTime(path, lastAccessTime.WrappedDateTime);
        }

        public  void SetLastAccessTimeUtc(string path, DateTime lastAccessTimeUtc)
        {
            GIO.File.SetLastAccessTimeUtc(path, lastAccessTimeUtc.WrappedDateTime);
        }

        public void SetLastWriteTime(string path, DateTime lastWriteTime)
        {
            GIO.File.SetLastWriteTime(path, lastWriteTime.WrappedDateTime);
        }

        public void SetLastWriteTimeUtc(string path, DateTime lastWriteTimeUtc)
        {
            GIO.File.SetLastWriteTimeUtc(path, lastWriteTimeUtc.WrappedDateTime);
        }

        public void WriteAllBytes(string path, [In] ref byte[] bytes)
        {
            GIO.File.WriteAllBytes(path, bytes);
        }

        public void WriteAllLines(string path, [In] ref string[] contents)
        {
            GIO.File.WriteAllLines(path, contents);
        }

        public void WriteAllLines(string path, [In] ref string[] contents, Encoding encoding)
        {
            GIO.File.WriteAllLines(path, contents, encoding.UnWrapEncoding());
        }

        public void WriteAllLines(string path, GCollections.IEnumerable contents)
        {
            GIO.File.WriteAllLines(path, (GCollections.Generic.IEnumerable<string>)contents);
        }

        public void WriteAllLines(string path, GCollections.IEnumerable contents, Encoding encoding)
        {
            GIO.File.WriteAllLines(path, (GCollections.Generic.IEnumerable<string>)contents, encoding.UnWrapEncoding());
        }

        public void WriteAllText(string path, string contents)
        {
            GIO.File.WriteAllText(path, contents);
        }

        public void WriteAllText(string path, string contents, Encoding encoding)
        {
            GIO.File.WriteAllText(path, contents, encoding.UnWrapEncoding());
        }


    }
}
