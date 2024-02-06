Attribute VB_Name = "RegexSplitExample4"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 6, 2024
'@LastModified February 6, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.split?view=netframework-4.8.1#system-text-regularexpressions-regex-split(system-string)

Option Explicit

''
' Starting with the .NET Framework 2.0, all captured text is also added to the
' returned array. For example, the following code uses two sets of capturing
' parentheses to extract the elements of a date, including the date delimiters,
' from a date string. The first set of capturing parentheses captures the hyphen,
' and the second set captures the forward slash.
''
Public Sub RegexSplitExample4()
    Dim pvtInput As String
    pvtInput = "07/14/2007"
    Dim pattern As String
    pattern = "(-)|(/)"
    Dim pvtRegex As DotNetLib.Regex
    Set pvtRegex = Regex.Create(pattern)
    Dim pvtResult As Variant
    For Each pvtResult In pvtRegex.Split(pvtInput)
        Debug.Print VBString.Format("'{0}'", pvtResult)
    Next
End Sub

' Output
'    '07'
'    '/'
'    '14'
'    '/'
'    '2007'

