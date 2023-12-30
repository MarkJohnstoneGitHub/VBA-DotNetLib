Attribute VB_Name = "StringIndexOfExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 30, 2023
'@LastModified December 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.indexof?view=netframework-4.8.1#system-string-indexof(system-string-system-int32-system-int32)

Option Explicit

''
' The following example finds the index of all occurrences of the string "he"
' within a substring of another string. Note that the number of characters to
' be searched must be recalculated for each search iteration.
''
Public Sub StringIndexOfExample()
    Dim br1 As String
    br1 = "0----+----1----+----2----+----3----+----4----+----5----+----6----+---"
    Dim br2 As String
    br2 = "012345678901234567890123456789012345678901234567890123456789012345678"
    Dim str As DotNetLib.String
    Set str = Strings.Create("Now is the time for all good men to come to the aid of their country.")
    Dim pvtStart As Long
    Dim at As Long
    Dim pvtEnd As Long
    Dim pvtCount As Long
    
    pvtEnd = str.length
    pvtStart = pvtEnd / 2
    Debug.Print
    Debug.Print VBAString.Format("All occurrences of 'he' from position {0} to {1}.", pvtStart, pvtEnd - 1)
    Debug.Print VBAString.Format("{1}{0}{2}{0}{3}{0}", Environment.NewLine, br1, br2, str);
    Debug.Print "The string 'he' occurs at position(s): ";
    
    pvtCount = 0
    at = 0
    
    Do While ((pvtStart <= pvtEnd) And (at > -1))
        ' start+count must be a position within -str-.
        pvtCount = pvtEnd - pvtStart
        at = str.IndexOf11("he", pvtStart, pvtCount)
        If (at = -1) Then
            Exit Do
        End If
        Debug.Print VBAString.Format("{0} ", at);
        pvtStart = at + 1
    Loop
End Sub

'/*
'This example produces the following results:
'
'All occurrences of 'he' from position 34 to 68.
'0----+----1----+----2----+----3----+----4----+----5----+----6----+---
'012345678901234567890123456789012345678901234567890123456789012345678
'Now is the time for all good men to come to the aid of their country.
'
'The string 'he' occurs at position(s): 45 56
'
'*/
