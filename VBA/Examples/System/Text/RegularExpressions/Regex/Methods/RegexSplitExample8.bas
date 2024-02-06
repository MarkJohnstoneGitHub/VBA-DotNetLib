Attribute VB_Name = "RegexSplitExample8"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 6, 2024
'@LastModified February 6, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.split?view=netframework-4.8.1#system-text-regularexpressions-regex-split(system-string-system-int32)

Option Explicit

''
' Starting with the .NET Framework 2.0, all captured text is added to the
' returned array. However, elements in the returned array that contain captured
' text are not counted in determining whether the number of matched substrings
' equals count. For example, in the following code, a regular expression uses
' two sets of capturing parentheses to extract the elements of a date from a
' date string. The first set of capturing parentheses captures the hyphen, and
' the second set captures the forward slash. The call to the Split(String, Int32)
' method then specifies a maximum of two elements in the returned array.
''
Public Sub RegexSplitExample8()
    Dim pvtInput As String
    pvtInput = "07/14/2007"
    Dim pattern As String
    pattern = "(-)|(/)"
    Dim pvtRegex As DotNetLib.Regex
    Set pvtRegex = Regex.Create(pattern)
    Dim pvtResult As Variant
    For Each pvtResult In pvtRegex.Split(pvtInput, 2)
        Debug.Print VBString.Format("'{0}'", pvtResult)
    Next
End Sub

' Output
'    '07'
'    '/'
'    '14/2007'
