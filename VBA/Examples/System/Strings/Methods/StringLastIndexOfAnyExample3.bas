Attribute VB_Name = "StringLastIndexOfAnyExample3"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 4, 2024
'@LastModified January 4, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.lastindexofany?view=netframework-4.8.1#system-string-lastindexofany(system-char()-system-int32-system-int32)

Option Explicit

''
' Sample for String.LastIndexOfAny(Char[], Int32, Int32)
' The following example finds the index of the last occurrence of any character
' in the string "aid" within a substring of another string.
''
Public Sub StringLastIndexOfAnyExample3()
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
       
    pvtStart = ((str.Length - 1) * 2) / 3
    pvtCount = (str.Length - 1) / 3
    Debug.Print VBString.Format("The last character occurrence from position {0} for {1} characters.", pvtStart, pvtCount)
    Debug.Print VBString.Format("{1}{0}{2}{0}{3}{0}", Environment.NewLine, br1, br2, str)
    Debug.Print VBString.Format("A character in '{0}' occurs at position: ", target);
    
    at = str.LastIndexOfAny3(anyOf, pvtStart, pvtCount)
    If (at > -1) Then
        Debug.Print at
    Else
        Debug.Print "(not found)"
    End If
    Debug.Print VBString.Format("{0}{0}{0}", Environment.NewLine)
End Sub

'/*
'This example produces the following results:
'The last character occurrence from position 44 for 22 characters.
'0----+----1----+----2----+----3----+----4----+----5----+----6----+-
'0123456789012345678901234567890123456789012345678901234567890123456
'Now is the time for all good men to come to the aid of their party.
'
'A character in 'aid' occurs at position: 27
'*/
