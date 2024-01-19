VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides a set of methods and properties that you can use to accurately measure elapsed time."
'@Folder "VBADotNetLib.System.Diagnostics"
'@PredeclaredId
'@ModuleDescription("Provides a set of methods and properties that you can use to accurately measure elapsed time.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 12, 2024
'@LastModified January 12, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.diagnostics.stopwatch?view=netframework-4.8.1

Option Explicit

Private Type TStopwatchSingleton
   StopwatchSingleton As DotNetLib.StopwatchSingleton
End Type

Private this As TStopwatchSingleton

Private Sub Class_Initialize()
   Set this.StopwatchSingleton = New DotNetLib.StopwatchSingleton
End Sub

Private Sub Class_Terminate()
   Set this.StopwatchSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the Stopwatch class.")
Public Function Create() As DotNetLib.Stopwatch
Attribute Create.VB_Description = "Initializes a new instance of the Stopwatch class."
   Set Create = this.StopwatchSingleton.Create()
End Function

'@Description("Gets the frequency of the timer as the number of ticks per second. This field is read-only.")
Public Property Get Frequency() As LongLong
Attribute Frequency.VB_Description = "Gets the frequency of the timer as the number of ticks per second. This field is read-only."
   Frequency = this.StopwatchSingleton.Frequency()
End Property

'@Description("Indicates whether the timer is based on a high-resolution performance counter. This field is read-only.")
Public Property Get IsHighResolution() As Boolean
Attribute IsHighResolution.VB_Description = "Indicates whether the timer is based on a high-resolution performance counter. This field is read-only."
   IsHighResolution = this.StopwatchSingleton.IsHighResolution()
End Property

'@Description("Gets the current number of ticks in the timer mechanism.")
Public Function GetTimestamp() As LongLong
Attribute GetTimestamp.VB_Description = "Gets the current number of ticks in the timer mechanism."
   GetTimestamp = this.StopwatchSingleton.GetTimestamp()
End Function

'@Description("Initializes a new Stopwatch instance, sets the elapsed time property to zero, and starts measuring elapsed time.")
Public Function StartNew() As DotNetLib.Stopwatch
Attribute StartNew.VB_Description = "Initializes a new Stopwatch instance, sets the elapsed time property to zero, and starts measuring elapsed time."
   Set StartNew = this.StopwatchSingleton.StartNew()
End Function
