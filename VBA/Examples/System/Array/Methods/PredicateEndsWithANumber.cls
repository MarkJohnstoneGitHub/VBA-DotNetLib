VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredicateEndsWithANumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Examples.System.Array.Methods")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 14, 2024
'@LastModified February 14, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.trueforall?view=netframework-4.8.1#examples

'@Remarks Used in example ArrayTrueForAllExample.bas

Option Explicit

Implements DotNetLib.IPredicate

''
' Determines whether the last character in a string is a number.
''
Public Function EndsWithANumber(ByVal pMatch As String) As Boolean
    Dim result As Long
    EndsWithANumber = Int32.TryParse(Right$(pMatch, 1), result)
End Function

'----------------------------------------------------------------------
' Interface IPredicate
'----------------------------------------------------------------------

Private Function IPredicate_CallBack(ByVal pMatch As Variant) As Boolean
    IPredicate_CallBack = Me.EndsWithANumber(pMatch)
End Function

