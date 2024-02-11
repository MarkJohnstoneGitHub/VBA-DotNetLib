VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredicateContainsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TestExamples.System.Array.Methods")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 10, 2024
'@LastModified February 11, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference

Option Explicit

Implements DotNetLib.IPredicate

Private Type TStringFilter
   pvtFilter As String
End Type

Private this As TStringFilter

Friend Property Get Self() As PredicateContainsString
   Set Self = Me
End Property

' Factory Methods

'Creates a predicate for a string filter, ignoring case
Public Function Create(ByVal pFilter As String) As DotNetLib.IPredicate
    With New PredicateContainsString
        .Filter = LCase$(pFilter)
        Set Create = Predicate.Create(.Self)
    End With
End Function

Public Property Get Filter() As String
   Filter = this.pvtFilter
End Property

Public Property Let Filter(ByVal pFilter As String)
   this.pvtFilter = pFilter
End Property

'----------------------------------------------------------------------
' Interface IPredicate
'----------------------------------------------------------------------

Private Function IPredicate_CallBack(ByVal pMatch As Variant) As Boolean
    IPredicate_CallBack = InStr(LCase$(pMatch), this.pvtFilter)
End Function
