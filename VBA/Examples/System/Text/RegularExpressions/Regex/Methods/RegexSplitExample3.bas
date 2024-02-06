Attribute VB_Name = "RegexSplitExample3"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 6, 2024
'@LastModified February 6, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.split?view=netframework-4.8.1#system-text-regularexpressions-regex-split(system-string)

Option Explicit

''
' If capturing parentheses are used in a Regex.Split expression, any captured
' text is included in the resulting string array. For example, if you split the
' string "plum-pear" on a hyphen placed within capturing parentheses, the
' returned array includes a string element that contains the hyphen.
''
Public Sub RegexSplitExample3()
    Dim pvtRegex As DotNetLib.Regex
    Set pvtRegex = Regex.Create("(-)")         ' Split on hyphens.
    Dim substrings() As String
    substrings = pvtRegex.Split("plum-pear")
    Dim pvtMatch As Variant
    For Each pvtMatch In substrings
        Debug.Print VBString.Format("'{0}'", pvtMatch)

    Next
End Sub

' The example displays the following output:
'    'plum'
'    '-'
'    'pear'
