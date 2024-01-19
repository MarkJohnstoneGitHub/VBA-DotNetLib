Attribute VB_Name = "StringLastIndexOfExample2"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 2, 2024
'@LastModified January 2, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.lastindexof?view=netframework-4.8.1#system-string-lastindexof(system-string-system-int32-system-int32-system-stringcomparison)

Option Explicit

''
' In the following example, the LastIndexOf(String, Int32, Int32, StringComparison)
' method is used to find the position of a soft hyphen (U+00AD) followed by an "m"
' in all but the first character position before the final "m" in two strings.
' Only one of the strings contains the required substring. If the example is run
' on .NET Framework 4 or later, in both cases, because the soft hyphen is an
' ignorable character, the method returns the index of "m" in the string when it
' performs a culture-sensitive comparison. When it performs an ordinal comparison,
' however, it finds the substring only in the first string. Note that in the case
' of the first string, which includes the soft hyphen followed by an "m", the
' method returns the index of the "m" when it performs a culture-sensitive comparison.
' The method returns the index of the soft hyphen in the first string only when it
' performs an ordinal comparison.
''
Public Sub StringLastIndexOfExample2()
    Dim searchString As DotNetLib.String
    Set searchString = Strings.CreateUnescape("\u00ADm")
    Dim s1 As DotNetLib.String
    Set s1 = Strings.CreateUnescape("ani\u00ADmal")
    Dim s2 As DotNetLib.String
    Set s2 = Strings.Create("animal")
    
    Dim position As Long
    position = s1.LastIndexOf7("m")
    If (position >= 1) Then
        Debug.Print s1.LastIndexOf6(searchString, position, position, StringComparison.StringComparison_CurrentCulture)
        Debug.Print s1.LastIndexOf6(searchString, position, position, StringComparison.StringComparison_Ordinal)
    End If
    
    position = s2.LastIndexOf7("m")
    If (position >= 1) Then
        Debug.Print s2.LastIndexOf6(searchString, position, position, StringComparison.StringComparison_CurrentCulture)
        Debug.Print s2.LastIndexOf6(searchString, position, position, StringComparison.StringComparison_Ordinal)
    End If
End Sub

' The example displays the following output:
'
' 4
' 3
' 3
' -1
