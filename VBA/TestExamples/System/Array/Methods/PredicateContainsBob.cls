VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredicateContainsBob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TestExamples.System.Array.Methods")
'@PredeclaredId
'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 8, 2024
'@LastModified February 11, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference

Option Explicit

Implements DotNetLib.IPredicate

'----------------------------------------------------------------------
' Interface IPredicate
'----------------------------------------------------------------------

'Private Function MyPredicate(ByVal pMatch As String) As Boolean
'    MyPredicate = InStr(LCase$(pMatch), "bob")
'End Function

''
'
' @Remarks Used for the
''
Private Function IPredicate_CallBack(ByVal pMatch As Variant) As Boolean
    IPredicate_CallBack = InStr(LCase$(pMatch), "bob")
End Function

