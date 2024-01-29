Attribute VB_Name = "RegexExamples"
'@IgnoreModule IndexedDefaultMemberAccess
'@Folder "Examples.System.Text.RegularExpressions.Regex"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 30, 2023
'@LastModified January 28, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?view=netframework-4.8.1

Option Explicit

' The following example uses a regular expression to check for repeated occurrences
' of words in a string. The regular expression \b(?<word>\w+)\s+(\k<word>)\b can be
' interpreted as shown in the following table.
' pattern       Description
' \b            Start the match at a word boundary.
' (?<word>\w+)  Match one or more word characters up to a word boundary. Name this captured group word.
' \s+           Match one or more white-space characters.
' (\k<word>)    Match the captured group that is named word.
' \b            Match a word boundary.
Public Sub RegexExample1()
    ' Define a regular expression for repeated words.
    Dim rx As DotNetLib.Regex
    Set rx = Regex.Create("\b(?<word>\w+)\s+(\k<word>)\b", RegexOptions.RegexOptions_Compiled Or RegexOptions.RegexOptions_IgnoreCase)

    ' Define a test string.
    Dim text As String
    text = "The the quick brown fox  fox jumps over the lazy dog dog."
    
    ' Find matches.
    Dim pvtMatches As DotNetLib.MatchCollection
    Set pvtMatches = rx.Matches(text)

    ' Report the number of matches found.
    Debug.Print VBString.Format(Regex.Unescape("{0} matches found in:\n   {1}"), pvtMatches.Count, text)

    ' Report on each match.
    Dim varMatch As Variant
    For Each varMatch In pvtMatches
        Dim pvtMatch As DotNetLib.Match
        Set pvtMatch = varMatch
        
        Dim Groups As DotNetLib.GroupCollection
        Set Groups = pvtMatch.Groups
        Debug.Print VBString.Format("'{0}' repeated at positions {1} and {2}", _
                                    Groups.Item_2("word").value, _
                                    Groups(0).index, _
                                    Groups(1).index)
    Next
End Sub

' The example produces the following output to the console:
'       3 matches found in:
'          The the quick brown fox  fox jumps over the lazy dog dog.
'       'The' repeated at positions 0 and 4
'       'fox' repeated at positions 20 and 25
'       'dog' repeated at positions 49 and 53


' https://stackoverflow.com/questions/77130177/greediness-in-vba-regex/77130622#77130622
Public Sub TestRegex()
    ' Define a regular expression for repeated words.
    Dim rx As DotNetLib.Regex
    Set rx = Regex.Create("\[\[\w{5,}[^[\]]*\]\]", RegexOptions.RegexOptions_Compiled Or RegexOptions.RegexOptions_IgnoreCase)

    ' Define a test string.
    Dim text As String
    text = "[[FOOTNOTE 23]] + [[FOOTNOTE 1787] +  182 - x +[[XYZ123]] + [[CD78G]]"
    
    ' Find matches.
    Dim pvtMatches As DotNetLib.MatchCollection
    Set pvtMatches = rx.Matches(text)

End Sub


