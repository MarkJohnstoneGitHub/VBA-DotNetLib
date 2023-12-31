Attribute VB_Name = "StringIndexOfAnyExample3"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 31, 2023
'@LastModified December 31, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.indexofany?view=netframework-4.8.1#system-string-indexofany(system-char()-system-int32-system-int32)

Option Explicit

''
' Sample for String.IndexOfAny(Char[], Int32, Int32)
' The following example finds the index of the occurrence of any character of
' the string "aid" within a substring of another string.
''
Public Sub StringIndexOfAnyExample3()
    Dim br1 As String
    br1 = "0----+----1----+----2----+----3----+----4----+----5----+----6----+-"
    Dim br2 As String
    br2 = "0123456789012345678901234567890123456789012345678901234567890123456"
    Dim str As DotNetLib.String
    Set str = Strings.Create("Now is the time for all good men to come to the aid of their party.")
    Dim pvtStart As Long
    Dim at As Long
    Dim pvtCount As Long
    
    Dim target As DotNetLib.String
    Set target = Strings.Create("aid")
    Dim anyOf As DotNetLib.String
    Set anyOf = target
    
    pvtStart = (str.length - 1) / 3
    pvtCount = (str.length - 1) / 4
    Debug.Print
    Debug.Print VBAString.Format("The first character occurrence from position {0} for {1} characters.", pvtStart, pvtCount)
    Debug.Print VBAString.Format("{1}{0}{2}{0}{3}{0}", Environment.NewLine, br1, br2, str)
    Debug.Print VBAString.Format("A character in '{0}' occurs at position: ", target);

    at = str.IndexOfAny3(anyOf, pvtStart, pvtCount)
    If (at > -1) Then
        Debug.Print at;
    Else
        Debug.Print "(not found)";
    End If
    Debug.Print
End Sub

'/*
'
'The first character occurrence from position 22 for 16 characters.
'0----+----1----+----2----+----3----+----4----+----5----+----6----+-
'0123456789012345678901234567890123456789012345678901234567890123456
'Now is the time for all good men to come to the aid of their party.
'
'A character in 'aid' occurs at position: 27
'
'*/
