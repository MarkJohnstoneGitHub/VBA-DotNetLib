Attribute VB_Name = "RegexSplitExample"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 6, 2024
'@LastModified February 6, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.split?view=netframework-4.8.1#system-text-regularexpressions-regex-split(system-string)

Option Explicit

''
' The Regex.Split methods are similar to the String.Split(Char[]) method, except
' that Regex.Split splits the string at a delimiter determined by a regular
' expression instead of a set of characters. The string is split as many times
' as possible. If no delimiter is found, the return value contains one element
' whose value is the original input string.
'
' If multiple matches are adjacent to one another, an empty string is inserted
' into the array. For example, splitting a string on a single hyphen causes the
' returned array to include an empty string in the position where two adjacent
' hyphens are found, as the following code shows.
''
Public Sub RegexSplitExample()
    Dim pvtRegex As DotNetLib.Regex
    Set pvtRegex = Regex.Create("-")         ' Split on hyphens.
    Dim substrings() As String
    substrings = pvtRegex.Split("plum--pear")
    Dim pvtMatch As Variant
    For Each pvtMatch In substrings
        Debug.Print VBString.Format("'{0}'", pvtMatch)
    Next
End Sub

' The example displays the following output:
'    'plum'
'    ''
'    'pear'
