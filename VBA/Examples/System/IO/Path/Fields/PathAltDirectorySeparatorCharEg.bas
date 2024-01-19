Attribute VB_Name = "PathAltDirectorySeparatorCharEg"
'@Folder "Examples.System.IO.Path.Fields"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.altdirectoryseparatorchar?view=netframework-4.8.1#examples

Option Explicit

Public Sub PathAltDirectorySeparatorChar()
    Debug.Print VBString.Format("Path.DirectorySeparatorChar: '{0}'", Path.DirectorySeparatorChar)
    Debug.Print VBString.Format("Path.AltDirectorySeparatorChar: '{0}'", Path.AltDirectorySeparatorChar)
    Debug.Print VBString.Format("Path.PathSeparator: '{0}'", Path.PathSeparator)
    Debug.Print VBString.Format("Path.VolumeSeparatorChar: '{0}'", Path.VolumeSeparatorChar)
    
    Dim invalidChars() As String
    invalidChars = Path.GetInvalidPathChars()
    Debug.Print ("Path.GetInvalidPathChars:")
    Dim ctr As Long
    For ctr = 0 To UBound(invalidChars)
        'TODO Implement UInt16, Convert.ToUInt16(String) for this example
        'Debug.Print VBAString.Format("  U+{0:X4)} ",Convert.ToUInt16(invalidChars[ctr]));
        If ((ctr + 1) Mod 10 = 0) Then
            Debug.Print
        End If
    Next
    Debug.Print
    
End Sub
'using System;
'using System.IO;
'
'Class Program
'{
'    static void Main()
'    {
'        Console.WriteLine($"Path.DirectorySeparatorChar: '{Path.DirectorySeparatorChar}'");
'        Console.WriteLine($"Path.AltDirectorySeparatorChar: '{Path.AltDirectorySeparatorChar}'");
'        Console.WriteLine($"Path.PathSeparator: '{Path.PathSeparator}'");
'        Console.WriteLine($"Path.VolumeSeparatorChar: '{Path.VolumeSeparatorChar}'");
'        var invalidChars = Path.GetInvalidPathChars();
'        Console.WriteLine($"Path.GetInvalidPathChars:");
'        for (int ctr = 0; ctr < invalidChars.Length; ctr++)
'        {
'            Console.Write($"  U+{Convert.ToUInt16(invalidChars[ctr]):X4} ");
'            if ((ctr + 1) % 10 == 0) Console.WriteLine();
'        }
'        Console.WriteLine();
'    }
'}
'// The example displays the following output when run on a Windows system:
'//    Path.DirectorySeparatorChar: '\'
'//    Path.AltDirectorySeparatorChar: '/'
'//    Path.PathSeparator: ';'
'//    Path.VolumeSeparatorChar: ':'
'//    Path.GetInvalidPathChars:
'//      U+007C)   U+0000)   U+0001)   U+0002)   U+0003)   U+0004)   U+0005)   U+0006)   U+0007)   U+0008)
'//      U+0009)   U+000A)   U+000B)   U+000C)   U+000D)   U+000E)   U+000F)   U+0010)   U+0011)   U+0012)
'//      U+0013)   U+0014)   U+0015)   U+0016)   U+0017)   U+0018)   U+0019)   U+001A)   U+001B)   U+001C)
'//      U+001D)   U+001E)   U+001F)
'//
'// The example displays the following output when run on a Linux system:
'//    Path.DirectorySeparatorChar: '/'
'//    Path.AltDirectorySeparatorChar: '/'
'//    Path.PathSeparator: ':'
'//    Path.VolumeSeparatorChar: '/'
'//    Path.GetInvalidPathChars:
'//      U+0000
