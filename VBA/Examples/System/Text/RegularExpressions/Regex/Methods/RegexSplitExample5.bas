Attribute VB_Name = "RegexSplitExample5"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 6, 2024
'@LastModified February 6, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.split?view=netframework-4.8.1#system-text-regularexpressions-regex-split(system-string)

Option Explicit

''
' If the regular expression can match the empty string, Split(String) will split
' the string into an array of single-character strings because the empty string
' delimiter can be found at every location. For example:
''
Public Sub RegexSplitExample5()
    Dim pvtInput As String
    pvtInput = "characters"
    Dim pvtRegex As DotNetLib.Regex
    Set pvtRegex = Regex.Create("")
    Dim substrings() As String
    substrings = pvtRegex.Split(pvtInput)
    Debug.Print "{";
    Dim ctr As Long
    For ctr = 0 To UBound(substrings)
        Debug.Print substrings(ctr);
        If (ctr < UBound(substrings)) Then
            Debug.Print ", ";
        End If
    Next
    Debug.Print "}"
End Sub

' The example displays the following output:
'    {, c, h, a, r, a, c, t, e, r, s, }

