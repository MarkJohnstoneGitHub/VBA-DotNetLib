Attribute VB_Name = "RegexSplitExample2"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 6, 2024
'@LastModified February 6, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.split?view=netframework-4.8.1#system-text-regularexpressions-regex-split(system-string)

Option Explicit

''
' If a match is found at the beginning or the end of the input string, an empty
' string is included at the beginning or the end of the returned array.
' The following example uses the regular expression pattern \d+ to split an
' input string on numeric characters. Because the string begins and ends with
' matching numeric characters, the value of the first and last element of the
' returned array is String.Empty.
''
Public Sub RegexSplitExample2()
    Dim pattern As String
    pattern = "\d+"
    Dim rgx As DotNetLib.Regex
    Set rgx = Regex.Create(pattern)
    Dim pvtInput As String
    pvtInput = "123ABCDE456FGHIJKL789MNOPQ012"
    
    Dim pvtResult() As String
    pvtResult = rgx.Split(pvtInput)
    Dim ctr As Long
    For ctr = 0 To UBound(pvtResult)
        Debug.Print VBString.Format("'{0}'", pvtResult(ctr));
        If (ctr < UBound(pvtResult)) Then
            Debug.Print (", ");
        End If
    Next
    Debug.Print
End Sub

' The example displays the following output:
'       '', 'ABCDE', 'FGHIJKL', 'MNOPQ', ''
