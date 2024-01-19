Attribute VB_Name = "StringIndexOfExample6"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 31, 2023
'@LastModified December 31, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.indexof?view=netframework-4.8.1#system-string-indexof(system-string)

Option Explicit

Public Sub StringIndexOfExample6()
    Dim s1 As DotNetLib.String
    Set s1 = Strings.CreateUnescape("ani\u00ADmal")
    Dim s2 As DotNetLib.String
    Set s2 = Strings.Create("animal")

    ' Find the index of the soft hyphen.
    Debug.Print s1.IndexOf(Strings.CreateUnescape("\u00AD"))
    Debug.Print s2.IndexOf(Strings.CreateUnescape("\u00AD"))
    
    ' Find the index of the soft hyphen followed by "n".
    Debug.Print s1.IndexOf(Strings.CreateUnescape("\u00ADn"))
    Debug.Print s2.IndexOf(Strings.CreateUnescape("\u00ADn"))

    ' Find the index of the soft hyphen followed by "m".
    Debug.Print s1.IndexOf(Strings.CreateUnescape("\u00ADm"))
    Debug.Print s2.IndexOf(Strings.CreateUnescape("\u00ADm"))
End Sub

' The example displays the following output
' if run under the .NET Framework 4 or later:
'       0
'       0
'       1
'       1
'       4
'       3
