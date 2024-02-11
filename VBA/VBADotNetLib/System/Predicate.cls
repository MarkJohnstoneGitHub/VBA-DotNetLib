VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Predicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents the method that defines a set of criteria and determines whether the specified object meets those criteria."
'@Folder "VBADotNetLib.System"
'@PredeclaredId
'@ModuleDescription("Represents the method that defines a set of criteria and determines whether the specified object meets those criteria.")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 10, 2024
'@LastModified February 10, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.predicate-1?view=netframework-4.8.1

Option Explicit

Private Type TPredicateSingleton
   PredicateSingleton As DotNetLib.PredicateSingleton
End Type

Private this As TPredicateSingleton

Private Sub Class_Initialize()
   Set this.PredicateSingleton = New DotNetLib.PredicateSingleton
End Sub

Private Sub Class_Terminate()
   Set this.PredicateSingleton = Nothing
End Sub

'@Description("Assigns the predicate method to the Predicate delegate.")
Public Function Create(ByVal pPredicate As DotNetLib.Predicate) As DotNetLib.Predicate
Attribute Create.VB_Description = "Assigns the predicate method to the Predicate delegate."
   Set Create = this.PredicateSingleton.Create(pPredicate)
End Function
