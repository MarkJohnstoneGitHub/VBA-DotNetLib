using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("570ECA2F-EE5B-4BB4-92D6-90DB113E5DA5")]
    [Description("Performs operations on String instances that contain file or directory path information. These operations are performed in a cross-platform manner.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IPathSingleton
    {
        string AltDirectorySeparatorChar
        {
            [Description("Provides a platform-specific alternate character used to separate directory levels in a path string that reflects a hierarchical file system organization.")]
            get;
        }

        string DirectorySeparatorChar
        {
            [Description("Provides a platform-specific character used to separate directory levels in a path string that reflects a hierarchical file system organization.")]
            get;
        }

        string PathSeparator
        {
            [Description("A platform-specific separator character used to separate path strings in environment variables.")]
            get;
        }

        string VolumeSeparatorChar
        {
            [Description("Provides a platform-specific volume separator character.")]
            get;
        }

        // Methods

        [Description("Changes the extension of a path string.")]
        string ChangeExtension(string path, string extension);

        [Description("Combines an array of strings into a path.")]
        string Combine([In] ref string[] paths);

        [Description("Combines two strings into a path.")]
        string Combine(string path1, string path2);

        [Description("Combines three strings into a path.")]
        string Combine(string path1, string path2, string path3);

        [Description("Combines four strings into a path.")]
        string Combine(string path1, string path2, string path3, string path4);

        [Description("Returns the directory information for the specified path.")]
        string GetDirectoryName(string path);

        [Description("Returns the extension (including the period .) of the specified path string.")]
        string GetExtension(string path);

        [Description("Returns the file name and extension of the specified path string.")]
        string GetFileName(string path);

        [Description("Returns the file name of the specified path string without the extension.")]
        string GetFileNameWithoutExtension(string path);

        [Description("Returns the absolute path for the specified path string.")]
        string GetFullPath(string path);

        [Description("Gets an array containing the characters that are not allowed in file names.")]
        string[] GetInvalidFileNameChars();

        [Description("Gets an array containing the characters that are not allowed in path names.")]
        string[] GetInvalidPathChars();

        [Description("Gets the root directory information from the path contained in the specified string.")]
        string GetPathRoot(string path);

        [Description("Returns a random folder name or file name.")]
        string GetRandomFileName();

        [Description("Creates a uniquely named, zero-byte temporary file on disk and returns the full path of that file.")]
        string GetTempFileName();

        [Description("Returns the path of the current user's temporary folder.")]
        string GetTempPath();

        [Description("Determines whether a path includes a file name extension.")]
        bool HasExtension(string path);

        [Description("Returns a value indicating whether the specified path string contains a root.")]
        bool IsPathRooted(string path);

    }
}
