VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatchEvaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents the method that is called each time a regular expression match is found during a Replace method operation."
'@Folder "VBADotNetLib.System.Text.RegularExpressions"
'@PredeclaredId
'@ModuleDescription("Represents the method that is called each time a regular expression match is found during a Replace method operation.")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 4, 2024
'@LastModified February 4, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.matchevaluator?view=netframework-4.8.1

Option Explicit

Private Type TMatchEvaluatorSingleton
   MatchEvaluatorSingleton As DotNetLib.MatchEvaluatorSingleton
End Type

Private this As TMatchEvaluatorSingleton

Private Sub Class_Initialize()
   Set this.MatchEvaluatorSingleton = New DotNetLib.MatchEvaluatorSingleton
End Sub

Private Sub Class_Terminate()
   Set this.MatchEvaluatorSingleton = Nothing
End Sub

'@Description("Assign the replace method to the MatchEvaluator delegate.")
Public Function Create(ByRef pMatchEvaluator As DotNetLib.IMatchEvaluator) As DotNetLib.MatchEvaluator
Attribute Create.VB_Description = "Assign the replace method to the MatchEvaluator delegate."
   Set Create = this.MatchEvaluatorSingleton.Create(pMatchEvaluator)
End Function
