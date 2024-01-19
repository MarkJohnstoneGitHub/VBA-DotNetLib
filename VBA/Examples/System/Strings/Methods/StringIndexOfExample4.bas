Attribute VB_Name = "StringIndexOfExample4"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 31, 2023
'@LastModified December 31, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.indexof?view=netframework-4.8.1#system-string-indexof(system-string-system-stringcomparison)

Option Explicit


Public Sub StringIndexOfExample4()
    Dim searchString As DotNetLib.String
    Set searchString = Strings.CreateUnescape("\u00ADm")
    Dim s1 As DotNetLib.String
    Set s1 = Strings.CreateUnescape("ani\u00ADmal")
    Dim s2 As DotNetLib.String
    Set s2 = Strings.Create("animal")

    Debug.Print s1.IndexOf4(searchString, 2, StringComparison.StringComparison_CurrentCulture)
    Debug.Print s1.IndexOf4(searchString, 2, StringComparison.StringComparison_Ordinal)
    Debug.Print s2.IndexOf4(searchString, 2, StringComparison.StringComparison_CurrentCulture)
    Debug.Print s2.IndexOf4(searchString, 2, StringComparison.StringComparison_Ordinal)
End Sub

' The example displays the following output:
'       4
'       3
'       3
'       -1
