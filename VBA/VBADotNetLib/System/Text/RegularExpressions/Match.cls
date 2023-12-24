VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Match"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents the results from a single regular expression match."
'@Folder "VBADotNetLib.System.Text.RegularExpressions"
'@PredeclaredId
'@ModuleDescription("Represents the results from a single regular expression match.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 1, 2023
'@LastModified December 24 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.match?view=netframework-4.8.1

'@Remarks
'Singleton class for Match containing static members.

Option Explicit

Private Type TMatchSingleton
   MatchSingleton As DotNetLib.MatchSingleton
End Type

Private this As TMatchSingleton

Private Sub Class_Initialize()
   Set this.MatchSingleton = New DotNetLib.MatchSingleton
End Sub

Private Sub Class_Terminate()
   Set this.MatchSingleton = Nothing
End Sub

''
'@Static
'@Description("Gets the empty group. All failed matches return this empty match.")
''
Public Property Get EmptyMatch() As DotNetLib.Match
Attribute EmptyMatch.VB_Description = "Gets the empty group. All failed matches return this empty match."
   Set EmptyMatch = this.MatchSingleton.[Empty]()
End Property

''
'@Static
'@Description("Returns a Match instance equivalent to the one supplied that is suitable to share between multiple threads.")
''
Public Function Synchronized(ByVal inner As DotNetLib.Match) As DotNetLib.Match
Attribute Synchronized.VB_Description = "Returns a Match instance equivalent to the one supplied that is suitable to share between multiple threads."
   Set Synchronized = this.MatchSingleton.Synchronized(inner)
End Function
