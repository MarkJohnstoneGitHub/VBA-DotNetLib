Attribute VB_Name = "RegexReplace3Example"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 5, 2024
'@LastModified February 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.replace?view=netframework-4.8.1#system-text-regularexpressions-regex-replace(system-string-system-string-system-int32-system-int32)

Option Explicit

''
' The following example double-spaces all but the first line of a string.
' It defines a regular expression pattern, ^.*$, that matches a line of text,
' calls the Match(String) method to match the first line of the string, and
' uses the Match.Index and Match.Count properties to determine the starting
' position of the second line.
''
Public Sub RegexReplace3Example()
    Dim pvtInput As String
    pvtInput = VBString.Unescape("Instantiating a New Type\n" & _
                "Generally, there are two ways that an\n" & _
                "instance of a class or structure can\n" & _
                "be instantiated. ")
    Dim pattern As String
    pattern = "^.*$"
    Dim replacement As String
    replacement = VBString.Unescape("\n$&")
    Dim rgx   As DotNetLib.Regex
    Set rgx = Regex.Create(pattern, RegexOptions.RegexOptions_Multiline)
    Dim pvtResult As String
    pvtResult = VBA.vbNullString

    Dim pvtMatch As DotNetLib.Match
    Set pvtMatch = rgx.Match(pvtInput)
    ' Double space all but the first line.
    If (pvtMatch.Success) Then
        pvtResult = rgx.Replace3(pvtInput, replacement, -1, pvtMatch.index + pvtMatch.Length + 1)
    End If

    Debug.Print pvtResult
End Sub

' The example displays the following output:
'       Instantiating a New Type
'
'       Generally, there are two ways that an
'
'       instance of a class or structure can
'
'       be instntiated.
