Attribute VB_Name = "CaptureCollectionExample"
'@Folder("Examples.System.Text.RegularExpressions.CaptureCollection")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 30, 2023
'@LastModified September 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.capturecollection?view=netframework-4.8.1

Option Explicit

' The following example compares the Capture objects in the CaptureCollection
' object returned by the Group.Captures and Match.Captures properties. It also
' compares Group objects with the Capture objects in the CaptureCollection
' returned by the Group.Captures property. The example uses the following two
' regular expressions to find matches in a single input string:
' \b\w+\W{1,2}
' This regular expression pattern identifies a word that consists of one or
' more word characters, followed by either one or two non-word characters
' such as white space or punctuation. The regular expression does not include
' any capturing groups. The output from the example shows that both the Match
' object and the CaptureCollection objects returned by the Group.Captures and
' Match.Captures properties contain information about the same match.
'
' (\b\w+\W{1,2})+
' This regular expression pattern identifies the words in a sentence.
' The pattern defines a single capturing group that consists of one or more
' word characters followed by one or two non-word characters. The regular
' expression pattern uses the + quantifier to match one or more occurrences
' of this group. The output from this example shows that the Match object
' and the CaptureCollection object returned by the Match.Captures property
' contain information about the same match. The second Group object, which
' corresponds to the only capturing group in the regular expression, identifies
' only the last captured string, whereas the CaptureCollection object returned
' by the first capturing group's Group.Captures property includes all captured
' substrings.

Public Sub CaptureCollection()
    Dim pattern As String
    Dim strInput As String
    strInput = "The young, hairy, and tall dog slowly walked across the yard."
    Dim pvtMatch As DotNetLib.Match
    
    ' Match a word with a pattern that has no capturing groups.
    pattern = "\b\w+\W{1,2}"
    Set pvtMatch = Regex.Match(strInput, pattern)
    Debug.Print "Pattern: " + pattern
    Debug.Print "Match: " + pvtMatch.value
    Debug.Print VBAString.Format("  Match.Captures: {0}", pvtMatch.Captures.Count)
    Dim ctr As Long
    For ctr = 0 To pvtMatch.Captures.Count - 1
        Debug.Print VBAString.Format("    {0}: '{1}'", ctr, pvtMatch.Captures(ctr).value)
    Next
    Debug.Print VBAString.Format("  Match.Groups: {0}", pvtMatch.Groups.Count)
    Dim groupCtr As Long
    For groupCtr = 0 To pvtMatch.Groups.Count - 1
        Debug.Print VBAString.Format("    Group {0}: '{1}'", groupCtr, pvtMatch.Groups(groupCtr).value)
        Debug.Print VBAString.Format("    Group({0}).Captures: {1}", groupCtr, pvtMatch.Groups(groupCtr).Captures.Count)
        
        Dim captureCtr As Long
        For captureCtr = 0 To pvtMatch.Groups(groupCtr).Captures.Count - 1
        Debug.Print VBAString.Format("      Capture {0}: '{1}'", captureCtr, pvtMatch.Groups(groupCtr).Captures(captureCtr).value)
        Next
    Next
    Debug.Print VBAString.Format(Regex.Unescape("-----\n"))
    
    ' Match a sentence with a pattern that has a quantifier that
    ' applies to the entire group.
    pattern = "(\b\w+\W{1,2})+"
    Set pvtMatch = Regex.Match(strInput, pattern)
    Debug.Print "Pattern: " + pattern
    Debug.Print "Match: " + pvtMatch.value
    Debug.Print VBAString.Format("  Match.Captures: {0}", pvtMatch.Captures.Count)
    
    For ctr = 0 To pvtMatch.Captures.Count - 1
        Debug.Print VBAString.Format("    {0}: '{1}'", ctr, pvtMatch.Captures(ctr).value)
    Next
    Debug.Print VBAString.Format("  Match.Groups: {0}", pvtMatch.Groups.Count)
    For groupCtr = 0 To pvtMatch.Groups.Count - 1
        Debug.Print VBAString.Format("    Group {0}: '{1}'", groupCtr, pvtMatch.Groups(groupCtr).value)
        Debug.Print VBAString.Format("    Group({0}).Captures: {1}", groupCtr, pvtMatch.Groups(groupCtr).Captures.Count)
        
        For captureCtr = 0 To pvtMatch.Groups(groupCtr).Captures.Count - 1
        Debug.Print VBAString.Format("      Capture {0}: '{1}'", captureCtr, pvtMatch.Groups(groupCtr).Captures(captureCtr).value)
        Next
    Next
End Sub

' The example displays the following output:
'    Pattern: \b\w+\W{1,2}
'    Match: The
'      Match.Captures: 1
'        0: 'The '
'      Match.Groups: 1
'        Group 0: 'The '
'        Group(0).Captures: 1
'          Capture 0: 'The '
'    -----
'
'    Pattern: (\b\w+\W{1,2})+
'    Match: The young, hairy, and tall dog slowly walked across the yard.
'      Match.Captures: 1
'        0: 'The young, hairy, and tall dog slowly walked across the yard.'
'      Match.Groups: 2
'        Group 0: 'The young, hairy, and tall dog slowly walked across the yard.'
'        Group(0).Captures: 1
'          Capture 0: 'The young, hairy, and tall dog slowly walked across the yard.'
'        Group 1: 'yard.'
'        Group(1).Captures: 11
'          Capture 0: 'The '
'          Capture 1: 'young, '
'          Capture 2: 'hairy, '
'          Capture 3: 'and '
'          Capture 4: 'tall '
'          Capture 5: 'dog '
'          Capture 6: 'slowly '
'          Capture 7: 'walked '
'          Capture 8: 'across '
'          Capture 9: 'the '
'          Capture 10: 'yard.'
