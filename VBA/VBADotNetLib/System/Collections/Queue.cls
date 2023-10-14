VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a first-in, first-out collection of objects."
'@Folder "VBADotNetLib.System.Collections"
'@PredeclaredId
'@ModuleDescription("Represents a first-in, first-out collection of objects.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 15, 2023
'@LastModified October 15, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.queue?view=netframework-4.8.1

'@Remarks
'Singleton class for Queue containing factory methods and static members.

Option Explicit

Private Type TQueueSingleton
   QueueSingleton As DotNetLib.QueueSingleton
End Type

Private this As TQueueSingleton

Private Sub Class_Initialize()
   Set this.QueueSingleton = New DotNetLib.QueueSingleton
End Sub

Private Sub Class_Terminate()
   Set this.QueueSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the Queue class that is empty, has the default or specified initial capacity, and uses the default or specified growth factor.")
Public Function Create(Optional ByVal capacity As Long = 32, Optional ByVal growFactor As Single = 2) As DotNetLib.Queue
   Set Create = this.QueueSingleton.Create(capacity, growFactor)
End Function

'@Description("Initializes a new instance of the Queue class that contains elements copied from the specified collection, has the same initial capacity as the number of elements copied, and uses the default growth factor.")
Public Function Create2(ByRef col As DotNetLib.ICollection) As DotNetLib.Queue
   Set Create2 = this.QueueSingleton.Create2(col)
End Function

'@Description("Returns a new Queue that wraps the original queue, and is thread safe.")
Public Function Synchronized(ByRef Queue As DotNetLib.Queue) As DotNetLib.Queue
   Set Synchronized = this.QueueSingleton.Synchronized(Queue)
End Function

