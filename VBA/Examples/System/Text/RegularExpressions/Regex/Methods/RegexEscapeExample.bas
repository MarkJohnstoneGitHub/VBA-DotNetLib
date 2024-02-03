Attribute VB_Name = "RegexEscapeExample"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 3, 2024
'@LastModified February 3, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.escape?view=netframework-4.8.1#examples

Option Explicit

''
' The following example extracts comments from text. It assumes that the
' comments are delimited by a begin comment symbol and an end comment symbol
' that is selected by the user. Because the comment symbols are to be
' interpreted literally, they are passed to the Escape method to ensure that
' they cannot be misinterpreted as metacharacters. In addition, the example
' explicitly checks whether the end comment symbol entered by the user is a
' closing bracket (]) or brace (}). If it is, a backslash character (\) is
' prepended to the bracket or brace so that it is interpreted literally.
' Note that the example also uses the Match.Groups collection to display the
' comment only, rather than the comment together with its opening and closing
' comment symbols.
''
Public Sub RegexEscapeExample()
    Dim beginComment As String
    beginComment = InputBox("Enter begin comment symbol: ")
    
    Dim endComment As String
    endComment = InputBox("Enter end comment symbol: ")
    
    Dim pvtInput As String
    pvtInput = "Text [comment comment comment] more text [comment]"
    Dim pattern As String
    pattern = Regex.Escape(beginComment) + "(.*?)"
    Dim endPattern As String
    endPattern = Regex.Escape(endComment)
    If endComment = "]" Or endComment = "}" Then
        endPattern = "\" + endPattern
    End If
    pattern = pattern + endPattern

    Dim pvtMatches As DotNetLib.MatchCollection
    Set pvtMatches = Regex.Matches(pvtInput, pattern)
    Debug.Print pattern
    Dim commentNumber As Long
    commentNumber = 0
    Dim varMatch As Variant
    For Each varMatch In pvtMatches
        Dim pvtMatch As DotNetLib.Match
        Set pvtMatch = varMatch
        commentNumber = commentNumber + 1
        Debug.Print VBString.Format("{0}: {1}", commentNumber, pvtMatch.Groups(1).value)
    Next
End Sub

' The example shows possible output from the example:
'       Enter begin comment symbol: [
'       Enter end comment symbol: ]
'       \[(.*?)\]
'       1: comment comment comment
'       2: comment
