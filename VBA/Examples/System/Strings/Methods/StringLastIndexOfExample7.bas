Attribute VB_Name = "StringLastIndexOfExample7"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 2, 2024
'@LastModified January 2, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.lastindexof?view=netframework-4.8.1#system-string-lastindexof(system-char-system-int32)

Option Explicit

''
' Sample for String.LastIndexOf(Char, Int32)
' The following example finds the index of all occurrences of a character in a
' string, working from the end of the string to the start of the string.
''
Public Sub StringLastIndexOfExample7()
    Dim br1 As String
    br1 = "0----+----1----+----2----+----3----+----4----+----5----+----6----+-"
    Dim br2 As String
    br2 = "0123456789012345678901234567890123456789012345678901234567890123456"
    Dim str As DotNetLib.String
    Set str = Strings.Create("Now is the time for all good men to come to the aid of their party.")
    Dim pvtStart As Long
    Dim at As Long
    
    pvtStart = str.length - 1
    Debug.Print
    Debug.Print VBAString.Format("All occurrences of 't' from position {0} to 0.", pvtStart)
    Debug.Print VBAString.Format("{1}{0}{2}{0}{3}{0}", Environment.NewLine, br1, br2, str);
    Debug.Print "The string 't' occurs at position(s): ";
    
    at = 0
    Do While ((pvtStart > -1) And (at > -1))
        at = str.LastIndexOf8("t", pvtStart)
        If (at > -1) Then
            Debug.Print VBAString.Format("{0} ", at);
            pvtStart = at - 1
        End If
    Loop
    Debug.Print VBAString.Format("{0}{0}{0}", Environment.NewLine)
End Sub

'/*
'This example produces the following results:
'All occurrences of 't' from position 66 to 0.
'0----+----1----+----2----+----3----+----4----+----5----+----6----+-
'0123456789012345678901234567890123456789012345678901234567890123456
'Now is the time for all good men to come to the aid of their party.
'
'The letter 't' occurs at position(s): 64 55 44 41 33 11 7
'*/
