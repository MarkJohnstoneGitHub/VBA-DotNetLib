Attribute VB_Name = "RegexGetGroupNamesExample"
'@Folder "Examples.System.Text.RegularExpressions.Regex.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 1, 2023
'@LastModified October 1, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.getgroupnames?view=netframework-4.8.1#examples

Option Explicit

''
' The following example defines a general-purpose ShowMatches method that
' displays the names of regular expression groups and their matched text.
''
Public Sub RegexGetGroupNames()
    Dim pattern As String
    pattern = "\b(?<FirstWord>\w+)\s?((\w+)\s)*(?<LastWord>\w+)?(?<Punctuation>\p{Po})"
    Dim strInput As String
    strInput = "The cow jumped over the moon."
    Dim rgx As DotNetLib.Regex
    Set rgx = Regex.Create(pattern)
    
    Dim pvtMatch As DotNetLib.Match
    Set pvtMatch = rgx.Match(strInput)
    If (pvtMatch.Success) Then
        ShowMatches rgx, pvtMatch
    End If

End Sub

Private Sub ShowMatches(ByVal r As DotNetLib.Regex, ByVal m As DotNetLib.Match)
    Dim names() As String
    names = r.GetGroupNames()
    Debug.Print "Named Groups:"
    
    Dim name As Variant
    For Each name In names
        Dim grp As DotNetLib.Group
        Set grp = m.Groups.Item_2(name)
        Debug.Print VBString.Format("   {0}: '{1}'", name, grp.value)
    Next
End Sub

' The example displays the following output:
'       Named Groups:
'          0: 'The cow jumped over the moon.'
'          1: 'the '
'          2: 'the'
'          FirstWord: 'The'
'          LastWord: 'moon'
'          Punctuation: '.'
