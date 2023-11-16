// https://learn.microsoft.com/en-us/dotnet/api/system.io.path?view=netframework-4.8.1

using GSystem = global::System;
using GIO = global::System.IO;
using System.Runtime.InteropServices;
using System;
using System.ComponentModel;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("Performs operations on String instances that contain file or directory path information. These operations are performed in a cross-platform manner.")]
    [Guid("87583304-B286-4FAD-BDEE-595A04F0B984")]
    [ProgId("DotNetLib.System.IO.PathSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IPathSingleton))]

    public class PathSingleton : IPathSingleton
    {
        public string AltDirectorySeparatorChar => GIO.Path.AltDirectorySeparatorChar.ToString();

        public string DirectorySeparatorChar => GIO.Path.DirectorySeparatorChar.ToString();

        public string PathSeparator => GIO.Path.PathSeparator.ToString();

        public string VolumeSeparatorChar => GIO.Path.VolumeSeparatorChar.ToString();

        // Methods

        public string ChangeExtension(string path, string extension)
        {
            return GIO.Path.ChangeExtension(path, extension);
        }

        public string Combine([In] ref string[] paths)
        {
            return GIO.Path.Combine(paths);
        }

        public string Combine(string path1, string path2)
        {
            return GIO.Path.Combine(path1, path2);
        }

        public string Combine(string path1, string path2, string path3)
        {
            return GIO.Path.Combine(path1, path2, path3);
        }

        public string Combine(string path1, string path2, string path3, string path4)
        {
            return GIO.Path.Combine(path1, path2, path3, path4);
        }

        public string GetDirectoryName(string path)
        {
            return GIO.Path.GetDirectoryName(path);
        }

        public string GetExtension(string path)
        {
            return GIO.Path.GetExtension(path);
        }

        public string GetFileName(string path)
        {  
            return GIO.Path.GetFileName(path);
        }

        public string GetFileNameWithoutExtension(string path)
        {
            return GIO.Path.GetFileNameWithoutExtension(path);
        }

        public string GetFullPath(string path)
        {
            return GIO.Path.GetFullPath(path);
        }

        // TODO Cache stringArr
        //TODO return array of string or string?
        public string[] GetInvalidFileNameChars()
        {
            char[] c = GIO.Path.GetInvalidFileNameChars();

            string[] stringArr = new string[c.Length]; //initialised
            int i = 0;
            foreach( char d in c)
            {
                stringArr[i] = char.ToString(d);
                i++;
            }
            return stringArr;
        }

        // TODO return array of string or string?
        // https://stackoverflow.com/questions/50873493/how-to-convert-character-array-char-to-string-array-string-without-loopi
        public string[] GetInvalidFileNameCharsV2()
        {
            char[] charArray = GIO.Path.GetInvalidFileNameChars();
            string[] result = new string(charArray).Split();
            return result;

            //char[] charArray = { 'A', 'B', 'C' }; // Character array initialized
            ///**
            //   *Below line will first convert a charArray to string using 
            //   *String(char[]) constructor and using String class method  
            //   *split(regularExpression) the converted string will
            //   *then be splited with empty string literal delimiter which in turn 
            //   *returns String[] 
            //   **/
            //String[] result = new String(charArray).split("");
        }

        // TODO return array of string or string?
        public GSystem.String[] GetInvalidPathChars()
        {
            char[] c = GIO.Path.GetInvalidPathChars();

            GSystem.String[] stringArr = new string[c.Length]; //initialised
            int i = 0;
            foreach (char d in c)
            {
                stringArr[i] = Char.ToString(d);
                i++;
            }
            return stringArr;
        }

        public string GetPathRoot(string path)
        {
            return GIO.Path.GetPathRoot(path);
        }

        public string GetRandomFileName()
        {
            return GIO.Path.GetRandomFileName();
        }

        public string GetTempFileName()
        {
            return GIO.Path.GetTempFileName();
        }

        public string GetTempPath()
        {
            return GIO.Path.GetTempPath();
        }

        public bool HasExtension(string path)
        {
            return GIO.Path.HasExtension(path);
        }

        public bool IsPathRooted(string path)
        {
            return GIO.Path.IsPathRooted(path);
        }
    }
}
