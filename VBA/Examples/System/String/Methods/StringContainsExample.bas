Attribute VB_Name = "StringContainsExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 29, 2023
'@LastModified December 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.contains?view=netframework-4.8.1

Option Explicit

''
' The following example determines whether the string "fox" is a substring of a
' familiar quotation. If "fox" is found in the string, it also displays its
' starting position.
''
Public Sub StringContainsExample()
    Dim s1 As DotNetLib.String
    Set s1 = Strings.Create("The quick brown fox jumps over the lazy dog")
    Dim s2 As DotNetLib.String
    Set s2 = Strings.Create("fox")
    Dim b As Boolean
    b = s1.Contains(s2)
    Debug.Print VBAString.Format("'{0}' is in the string '{1}': {2}", _
                s2, s1, b)
    If (b) Then
        Dim pvtIndex As Long
        pvtIndex = s1.IndexOf(s2)
        If (pvtIndex >= 0) Then
            Debug.Print VBAString.Format("'{0} begins at character position {1}", _
                              s2, pvtIndex + 1)
        End If
    End If
End Sub

' This example displays the following output:
'    'fox' is in the string 'The quick brown fox jumps over the lazy dog': True
'    'fox begins at character position 17
