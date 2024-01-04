Attribute VB_Name = "StringRemoveExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 4, 2024
'@LastModified January 4, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.remove?view=netframework-4.8.1#system-string-remove(system-int32)

Option Explicit

''
' This example demonstrates the String.Remove() method.
' The following example demonstrates the Remove method. The next-to-last case
' removes all text starting from the specified index through the end of the string.
' The last case removes three characters starting from the specified index.
''
Public Sub StringRemoveExample()
    Dim s As DotNetLib.String
    Set s = Strings.Create("abc---def")

    Debug.Print "Index: 012345678"
    Debug.Print VBAString.Format("1)     {0}", s)
    Debug.Print VBAString.Format("2)     {0}", s.Remove(3))
    Debug.Print VBAString.Format("3)     {0}", s.Remove2(3, 3))
End Sub

'/*
'This example produces the following results:
'
'Index: 012345678
'1)     abc---def
'2)     abc
'3)     abcdef
'
'*/
