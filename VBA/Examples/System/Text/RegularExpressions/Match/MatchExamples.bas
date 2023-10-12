Attribute VB_Name = "MatchExamples"
'@Folder("Examples.System.Text.RegularExpressions.Match")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 30, 2023
'@LastModified September 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.match?view=netframework-4.8.1

Option Explicit

''
' The following example calls the Regex.Matches(String, String) method to
' retrieve all pattern matches in an input string. It then iterates the
' Match objects in the returned MatchCollection object to display
' information about each match.
''
Public Sub MatchExample1()
    Dim strInput As String
    strInput = Regex.Unescape("int[] values = { 1, 2, 3 };\n" + _
                "for (int ctr = values.GetLowerBound(1); ctr <= values.GetUpperBound(1); ctr++)\n" + _
                "{\n" + _
                "   Console.Write(values[ctr]);\n" + _
                "   if (ctr < values.GetUpperBound(1))\n" + _
                "      Console.Write(\"", \"");\n" + _
                "}\n" + _
                "Console.WriteLine();\n")
    
    Dim pattern As String
    pattern = "Console\.Write(Line)?"
    
    Dim pvtMatches As DotNetLib.MatchCollection
    Set pvtMatches = Regex.Matches(strInput, pattern)
    
    Dim varMatch As Variant
    For Each varMatch In pvtMatches
        Dim pvtMatch As DotNetLib.Match
        Set pvtMatch = varMatch
        Debug.Print Strings.Format("'{0}' found in the source code at position {1}.", _
                                    pvtMatch.value, pvtMatch.index)
    Next
End Sub

' The example displays the following output:
'    'Console.Write' found in the source code at position 112.
'    'Console.Write' found in the source code at position 184.
'    'Console.WriteLine' found in the source code at position 207.

''
' The following example calls the Match(String, String) and NextMatch methods
' to retrieve one match at a time.
''
Public Sub MatchExample2()
    Dim strInput As String
    strInput = Regex.Unescape("int[] values = { 1, 2, 3 };\n" + _
                "for (int ctr = values.GetLowerBound(1); ctr <= values.GetUpperBound(1); ctr++)\n" + _
                "{\n" + _
                "   Console.Write(values[ctr]);\n" + _
                "   if (ctr < values.GetUpperBound(1))\n" + _
                "      Console.Write(\"", \"");\n" + _
                "}\n" + _
                "Console.WriteLine();\n")
                
    Dim pattern As String
    pattern = "Console\.Write(Line)?"
    Dim pvtMatch As DotNetLib.Match
    Set pvtMatch = Regex.Match(strInput, pattern)
    Do While (pvtMatch.Success)
        Debug.Print Strings.Format("'{0}' found in the source code at position {1}.", _
                                    pvtMatch.value, pvtMatch.index)
        Set pvtMatch = pvtMatch.NextMatch()
    Loop
End Sub

' The example displays the following output:
'    'Console.Write' found in the source code at position 112.
'    'Console.Write' found in the source code at position 184.
'    'Console.WriteLine' found in the source code at position 207.


''
'The Match object is immutable and has no public constructor. An instance of
' the Match class is returned by the Regex.Match method and represents the
' first pattern match in a string. Subsequent matches are represented by Match
' objects returned by the Match.NextMatch method. In addition, a MatchCollection
' object that consists of zero, one, or more Match objects is returned by the
' Regex.Matches method.
'
' If the Regex.Matches method fails to match a regular expression pattern in
' an input string, it returns an empty MatchCollection object. You can then use
' a foreach construct in C# or a For Each construct in Visual Basic to iterate
' the collection.
'
' If the Regex.Match method fails to match the regular expression pattern,
' it returns a Match object that is equal to Match.Empty. You can use the Success
' property to determine whether the match was successful.
' The following example provides an illustration.
''
Public Sub MatchExample3()
    ' Search for a pattern that is not found in the input string.
    Dim pattern As String
    pattern = "dog"
    Dim strInput As String
    strInput = "The cat saw the other cats playing in the back yard."
    Dim pvtMatch As DotNetLib.Match
    Set pvtMatch = Regex.Match(strInput, pattern)
    If (pvtMatch.Success) Then
        ' Report position as a one-based integer.
        Debug.Print Strings.Format("'{0}' was found at position {1} in '{2}'.", _
                                pvtMatch.value, pvtMatch.index + 1, strInput)
    Else
        Debug.Print Strings.Format("The pattern '{0}' was not found in '{1}'.", _
                                pattern, strInput)
    End If
    
End Sub

' The example displays the following output:
'     The pattern 'dog' was not found in 'The cat saw the other cats playing in the back yard.'.


