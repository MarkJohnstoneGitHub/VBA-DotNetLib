Attribute VB_Name = "CaptureExample"
'@Folder("Examples.System.Text.RegularExpressions.Capture")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 7, 2024
'@LastModified February 7, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.capture?view=netframework-4.8.1#examples

Option Explicit

''
' The following example defines a regular expression that matches sentences
' that contain no punctuation except for a period (".").
''
Public Sub CaptureExample()
    Dim pvtInput As String
    pvtInput = "Yes. This dog is very friendly."
    Dim pattern As String
    pattern = "((\w+)[\s.])+"
    Dim varMatch As Variant
    For Each varMatch In Regex.Matches(pvtInput, pattern)
        Dim pvtMatch As DotNetLib.Match
        Set pvtMatch = varMatch
        Debug.Print VBString.Format("Match: {0}", pvtMatch.value)
        Dim groupCtr As Long
        For groupCtr = 0 To pvtMatch.Groups.Count - 1
            Dim pvtGroup As DotNetLib.Group
            Set pvtGroup = pvtMatch.Groups(groupCtr)
            Debug.Print VBString.Format("   Group {0}: {1}", groupCtr, pvtGroup.value)
            Dim captureCtr As Long
            For captureCtr = 0 To pvtGroup.Captures.Count - 1
                Debug.Print VBString.Format("      Capture {0}: {1}", captureCtr, _
                                 pvtGroup.Captures(captureCtr).value)
            Next
        Next
    Next
End Sub

' The example displays the following output:
'       Match: Yes.
'          Group 0: Yes.
'             Capture 0: Yes.
'          Group 1: Yes.
'             Capture 0: Yes.
'          Group 2: Yes
'             Capture 0: Yes
'       Match: This dog is very friendly.
'          Group 0: This dog is very friendly.
'             Capture 0: This dog is very friendly.
'          Group 1: friendly.
'             Capture 0: This
'             Capture 1: dog
'             Capture 2: is
'             Capture 3: very
'             Capture 4: friendly.
'          Group 2: friendly
'             Capture 0: This
'             Capture 1: dog
'             Capture 2: is
'             Capture 3: very
'             Capture 4: friendly
