Attribute VB_Name = "RegexSplitExample6"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 6, 2024
'@LastModified February 6, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.split?view=netframework-4.8.1#system-text-regularexpressions-regex-split(system-string-system-int32)

Option Explicit

''
' In the following example, the regular expression /d+ is used to split an input
' string that includes one or more decimal digits into a maximum of three
' substrings. Because the beginning of the input string matches the regular
' expression pattern, the first array element contains String.Empty, the second
' contains the first set of alphabetic characters in the input string, and the
' third contains the remainder of the string that follows the third match.
''
Public Sub RegexSplitExample6()
    Dim pattern As String
    pattern = "\d+"
    Dim rgx As DotNetLib.Regex
    Set rgx = Regex.Create(pattern)
    Dim pvtInput As String
    pvtInput = "123ABCDE456FGHIJKL789MNOPQ012"

    Dim pvtResult() As String
    pvtResult = rgx.Split(pvtInput, 3)
    Dim ctr As Long
    For ctr = 0 To UBound(pvtResult)
        Debug.Print VBString.Format("'{0}'", pvtResult(ctr));
        If (ctr < UBound(pvtResult)) Then
            Debug.Print ", ";
        End If
    Next
    Debug.Print
End Sub

' The example displays the following output:
'       '', 'ABCDE', 'FGHIJKL789MNOPQ012'
