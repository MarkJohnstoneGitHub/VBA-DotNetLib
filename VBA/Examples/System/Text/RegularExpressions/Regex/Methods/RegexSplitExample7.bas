Attribute VB_Name = "RegexSplitExample7"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 6, 2024
'@LastModified February 6, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.split?view=netframework-4.8.1#system-text-regularexpressions-regex-split(system-string-system-int32)

Option Explicit

''
' If capturing parentheses are used in a regular expression, any captured text
' is included in the array of split strings. However, any array elements that
' contain captured text are not counted in determining whether the number of
' matches has reached count. For example, splitting the string
' "apple-apricot-plum-pear-banana" into a maximum of four substrings results in
' a seven-element array, as the following code shows.
''
Public Sub RegexSplitExample7()
    Dim pattern As String
    pattern = "(-)"
    Dim pvtInput As String
    pvtInput = "apple-apricot-plum-pear-banana"
    Dim pvtRegex As DotNetLib.Regex
    Set pvtRegex = Regex.Create(pattern)         ' Split on hyphens.
    Dim substrings() As String
    substrings = pvtRegex.Split(pvtInput, 4)
    Dim pvtMatch As Variant
    For Each pvtMatch In substrings
        Debug.Print VBString.Format("'{0}'", pvtMatch)
    Next
End Sub

' The example displays the following output:
'       'apple'
'       '-'
'       'apricot'
'       '-'
'       'plum'
'       '-'
'       'pear-banana'
