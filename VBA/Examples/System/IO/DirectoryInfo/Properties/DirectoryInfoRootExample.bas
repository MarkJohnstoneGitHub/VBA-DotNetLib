Attribute VB_Name = "DirectoryInfoRootExample"
'@Folder("Examples.System.IO.DirectoryInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 23, 2023
'@LastModified December 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.root?view=netframework-4.8.1#examples

Option Explicit

''
' The following example displays root locations for specified directories.
''
Public Sub DirectoryInfoRootExample()
    Dim di1 As DotNetLib.DirectoryInfo
    Set di1 = DirectoryInfo.Create("\\tempshare\tempdir")
    Dim di2 As DotNetLib.DirectoryInfo
    Set di2 = DirectoryInfo.Create("tempdir")
    Dim di3 As DotNetLib.DirectoryInfo
    Set di3 = DirectoryInfo.Create("x:\tempdir")
    Dim di4 As DotNetLib.DirectoryInfo
    Set di4 = DirectoryInfo.Create("c:\")
    
    Debug.Print VBAString.Format("The root path of '{0}' is '{1}'", di1.FullName, di1.Root)
    Debug.Print VBAString.Format("The root path of '{0}' is '{1}'", di2.FullName, di2.Root)
    Debug.Print VBAString.Format("The root path of '{0}' is '{1}'", di3.FullName, di3.Root)
    Debug.Print VBAString.Format("The root path of '{0}' is '{1}'", di4.FullName, di4.Root)
End Sub

'/*
'This code produces output similar to the following:
'
'The root path of '\\tempshare\tempdir' is '\\tempshare\tempdir'
'The root path of 'c:\Projects\ConsoleApplication1\ConsoleApplication1\bin\Debug\tempdir' is 'c:\'
'The root path of 'x:\tempdir' is 'x:\'
'The root path of 'c:\' is 'c:\'
'Press any key to continue . . .
'
'*/
