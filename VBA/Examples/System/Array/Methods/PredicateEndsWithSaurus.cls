VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredicateEndsWithSaurus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("Examples.System.Array.Methods")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 14, 2024
'@LastModified February 14, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.findlastindex?view=netframework-4.8.1#examples

'@Remarks Used in example ArrayFindLastIndexExample.bas

Option Explicit

Implements DotNetLib.IPredicate

''
' Search predicate returns true if a string ends in "saurus".
''
Public Function EndsWithSaurus(ByVal pMatch As String) As Boolean
    If (Len(pMatch) > 5) And LCase$(Right$(pMatch, 6)) = "saurus" Then
        EndsWithSaurus = True
    Else
        EndsWithSaurus = False
    End If
End Function

'----------------------------------------------------------------------
' Interface IPredicate
'----------------------------------------------------------------------

Private Function IPredicate_CallBack(ByVal pMatch As Variant) As Boolean
    IPredicate_CallBack = Me.EndsWithSaurus(pMatch)
End Function

