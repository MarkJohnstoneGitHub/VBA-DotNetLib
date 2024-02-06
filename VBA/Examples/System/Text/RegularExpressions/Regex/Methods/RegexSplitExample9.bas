Attribute VB_Name = "RegexSplitExample9"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 7, 2024
'@LastModified February 7, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.split?view=netframework-4.8.1#system-text-regularexpressions-regex-split(system-string-system-int32)

Option Explicit

''
' If the regular expression can match the empty string, Split(String, Int32)
' will split the string into an array of single-character strings because the
' empty string delimiter can be found at every location. The following example
' splits the string "characters" into as many elements as there are in the
' input string. Because the null string matches the beginning of the input
' string, a null string is inserted at the beginning of the returned array.
' This causes the tenth element to consist of the two characters at the end of
' the input string.
''
Public Sub RegexSplitExample9()
    Dim pvtInput As String
    pvtInput = "characters"
    Dim pvtRegex As DotNetLib.Regex
    Set pvtRegex = Regex.Create("")
    Dim substrings() As String
    substrings = pvtRegex.Split(pvtInput, Len(pvtInput))
    Debug.Print "{";
    Dim ctr As Long
    For ctr = 0 To UBound(substrings)
        Debug.Print substrings(ctr);
        If ctr < UBound(substrings) Then
            Debug.Print ", ";
        End If
    Next
    Debug.Print "}"
End Sub

' The example displays the following output:
'    {, c, h, a, r, a, c, t, e, rs}
