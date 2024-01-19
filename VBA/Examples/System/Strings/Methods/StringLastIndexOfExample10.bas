Attribute VB_Name = "StringLastIndexOfExample10"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 3, 2024
'@LastModified January 3, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.lastindexof?view=netframework-4.8.1#system-string-lastindexof(system-string-system-stringcomparison)

Option Explicit

''
' In the following example, the LastIndexOf(String, StringComparison) method is
' used to find two substrings (a soft hyphen followed by "n", and a soft hyphen
' followed by "m") in two strings. Only one of the strings contains a soft hyphen.
' If the example is run on .NET Framework 4 or later, because the soft hyphen is
' an ignorable character, a culture-sensitive search returns the same value that
' it would return if the soft hyphen were not included in the search string.
' An ordinal search, however, successfully finds the soft hyphen in one string
' and reports that it is absent from the second string.
''
Public Sub StringLastIndexOfExample10()
    Dim s1 As DotNetLib.String
    Set s1 = Strings.CreateUnescape("ani\u00ADmal")
    Dim s2 As DotNetLib.String
    Set s2 = Strings.Create("animal")
    
    Debug.Print "Culture-sensitive comparison:"
    
    ' Use culture-sensitive comparison to find the last soft hyphen followed by "n".
    Debug.Print s1.LastIndexOf4(Strings.CreateUnescape("\u00ADn"), StringComparison.StringComparison_CurrentCulture)
    Debug.Print s2.LastIndexOf4(Strings.CreateUnescape("\u00ADn"), StringComparison.StringComparison_CurrentCulture)

    ' Use culture-sensitive comparison to find the last soft hyphen followed by "m".
    Debug.Print s1.LastIndexOf4(Strings.CreateUnescape("\u00ADm"), StringComparison.StringComparison_CurrentCulture)
    Debug.Print s2.LastIndexOf4(Strings.CreateUnescape("\u00ADm"), StringComparison.StringComparison_CurrentCulture)

    Debug.Print "Ordinal comparison:"

    ' Use ordinal comparison to find the last soft hyphen followed by "n".
    Debug.Print s1.LastIndexOf4(Strings.CreateUnescape("\u00ADn"), StringComparison.StringComparison_Ordinal)
    Debug.Print s2.LastIndexOf4(Strings.CreateUnescape("\u00ADn"), StringComparison.StringComparison_Ordinal)

    ' Use ordinal comparison to find the last soft hyphen followed by "m".
    Debug.Print s1.LastIndexOf4(Strings.CreateUnescape("\u00ADm"), StringComparison.StringComparison_Ordinal)
    Debug.Print s2.LastIndexOf4(Strings.CreateUnescape("\u00ADm"), StringComparison.StringComparison_Ordinal)
End Sub

' The example displays the following output:
'
' Culture-sensitive comparison:
'  1
'  1
'  4
'  3
' Ordinal comparison:
' -1
' -1
'  3
' -1
