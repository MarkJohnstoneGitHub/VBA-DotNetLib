Attribute VB_Name = "StringIndexOfExample2"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 31, 2023
'@LastModified December 31, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.indexof?view=netframework-4.8.1#system-string-indexof(system-string-system-int32-system-int32)

Option Explicit

Public Sub StringIndexOfExample2()
    Dim searchString As DotNetLib.String
    Set searchString = Strings.CreateUnescape("\u00ADm")
    Dim s1 As DotNetLib.String
    Set s1 = Strings.CreateUnescape("ani\u00ADmal")
    Dim s2 As DotNetLib.String
    Set s2 = Strings.Create("animal")
    
    Debug.Print VBAString.Format(s1.IndexOf5(searchString, 2, 4))
    Debug.Print VBAString.Format(s2.IndexOf5(searchString, 2, 4))
End Sub

' The example displays the following output:
'       4
'       3
