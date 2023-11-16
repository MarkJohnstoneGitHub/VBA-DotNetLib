// https://learn.microsoft.com/en-us/dotnet/api/system.io.filesysteminfo?view=netframework-4.8.1
// Provides the base class for both FileInfo and DirectoryInfo objects.


using GSerialization = global::System.Runtime.Serialization;
using GIO = global::System.IO;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("CC37067A-BAF2-44AC-9F15-695F9F9C80ED")]
    [Description("Provides the base class for both FileInfo and DirectoryInfo objects.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface FileSystemInfo
    {
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

        bool Exists
        {
            [Description("Gets a value indicating whether the file or directory exists.")]
            get;
        }

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

        string Name
        {
            [Description("For files, gets the name of the file. For directories, gets the name of the last directory in the hierarchy if a hierarchy exists. Otherwise, the Name property gets the name of the directory.")]
            get;
        }

        // Methods
        [Description("Deletes a file or directory.")]
        void Delete();

        [Description("Sets the SerializationInfo object with the file name and additional exception information.")]
        void GetObjectData(GSerialization.SerializationInfo info, GSerialization.StreamingContext context);

        [Description("Refreshes the state of the object.")]
        void Refresh();
    }
}
