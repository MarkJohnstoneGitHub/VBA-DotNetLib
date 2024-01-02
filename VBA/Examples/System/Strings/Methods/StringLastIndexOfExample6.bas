Attribute VB_Name = "StringLastIndexOfExample6"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 2, 2024
'@LastModified January 2, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.lastindexof?view=netframework-4.8.1#system-string-lastindexof(system-string-system-int32)

Option Explicit

''
' Sample for String.LastIndexOf(String, Int32)
' The following example finds the index of all occurrences of a string in
' target string, working from the end of the target string to the start of
' the target string.
''
Public Sub StringLastIndexOfExample6()
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
    Debug.Print VBAString.Format("All occurrences of 'he' from position {0} to 0.", pvtStart)
    Debug.Print VBAString.Format("{1}{0}{2}{0}{3}{0}", Environment.NewLine, br1, br2, str)
    Debug.Print "The string 'he' occurs at position(s): ";
    
    at = 0
    Do While ((pvtStart > -1) And (at > -1))
        at = str.LastIndexOf8("he", pvtStart)
        If (at > -1) Then
            Debug.Print VBAString.Format("{0} ", at);
            pvtStart = at - 1
        End If
    Loop
End Sub

'/*
'This example produces the following results:
'All occurrences of 'he' from position 66 to 0.
'0----+----1----+----2----+----3----+----4----+----5----+----6----+-
'0123456789012345678901234567890123456789012345678901234567890123456
'Now is the time for all good men to come to the aid of their party.
'
'The string 'he' occurs at position(s): 56 45 8
'
'
'*/
