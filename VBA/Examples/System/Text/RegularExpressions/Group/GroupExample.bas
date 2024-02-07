Attribute VB_Name = "GroupExample"
'@Folder("Examples.System.Text.RegularExpressions.Group")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 7, 2024
'@LastModified February 7, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.group?view=netframework-4.8.1#remarks

' https://learn.microsoft.com/en-us/dotnet/standard/base-types/quantifiers-in-regular-expressions

Option Explicit

''
' A capturing group can capture zero, one, or more strings in a single match because of quantifiers.
' (For more information, see Quantifiers.) All the substrings matched by a single capturing group
' are available from the Group.Captures property. Information about the last substring captured can
' be accessed directly from the Value and Index properties. (That is, the Group instance is equivalent
' to the last item of the collection returned by the Captures property, which reflects the last
' capture made by the capturing group.)
'
' An example helps to clarify this relationship between a Group object and the
' System.Text.RegularExpressions.CaptureCollection that is returned by the Captures property.
' The regular expression pattern (\b(\w+?)[,:;]?\s?)+[?.!]matches entire sentences. The regular
' expression is defined as shown in the following table.
'
'   Pttern                  Description
'   \b                      Begin the match at a word boundary.
'   (\w+?)                  Match one or more word characters, but as few characters as possible.
'                           This is the second (inner) capturing group. (The first capturing group
'                           includes the \b language element.)
'   [,:;]?                  Match zero or one occurrence of a comma, colon, or semicolon.
'   \s?                     Match zero or one occurrence of a white-space character.
'   (\b(\w+?)[,:;]?\s?)+    Match the pattern consisting of a word boundary, one or more word characters,
'                           a punctuation symbol, and a white-space character one or more times.
'                           This is the first capturing group.
'   [?.!]   Match any occurrence of a period, question mark, or exclamation point.
'
' In this regular expression pattern, the subpattern (\w+?) is designed to match multiple words within
' a sentence. However, the value of the Group object represents only the last match that (\w+?)
' captures, whereas the Captures property returns a CaptureCollection that represents all captured text.
' As the output shows, the CaptureCollection for the second capturing group contains four objects.
' The last of these corresponds to the Group object.
''
Public Sub GroupExample()
    Dim pattern As String
    pattern = "(\b(\w+?)[,:;]?\s?)+[?.!]"
    Dim pvtInput As String
    pvtInput = "This is one sentence. This is a second sentence."
    
    Dim pvtMatch As DotNetLib.Match
    Set pvtMatch = Regex.Match(pvtInput, pattern)
    Debug.Print "Match: " & pvtMatch.value
    Dim groupCtr  As Long
    groupCtr = 0
    
    Dim varGroup As Variant
    For Each varGroup In pvtMatch.Groups
        Dim pvtGroup As DotNetLib.Group
        Set pvtGroup = varGroup
        groupCtr = groupCtr + 1
        Debug.Print VBString.Format("   Group {0}: '{1}'", groupCtr, pvtGroup.value)
        Dim captureCtr As Long
        captureCtr = 0
        Dim varCapture As Variant
        For Each varCapture In pvtGroup.Captures
            Dim pvtCapture As DotNetLib.Capture
            Set pvtCapture = varCapture
            captureCtr = captureCtr + 1
            Debug.Print VBString.Format("      Capture {0}: '{1}'", captureCtr, pvtCapture.value)
        Next
    Next
End Sub

' The example displays the following output:
'       Match: This is one sentence.
'          Group 1: 'This is one sentence.'
'             Capture 1: 'This is one sentence.'
'          Group 2: 'sentence'
'             Capture 1: 'This '
'             Capture 2: 'is '
'             Capture 3: 'one '
'             Capture 4: 'sentence'
'          Group 3: 'sentence'
'             Capture 1: 'This'
'             Capture 2: 'is'
'             Capture 3: 'one'
'             Capture 4: 'sentence'
