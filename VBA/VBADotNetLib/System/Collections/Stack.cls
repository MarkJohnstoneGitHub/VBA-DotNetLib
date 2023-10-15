VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a simple last-in-first-out (LIFO) non-generic collection of objects."
'@Folder "VBADotNetLib.System.Collections"
'@ModuleDescription("Represents a simple last-in-first-out (LIFO) non-generic collection of objects.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 15, 2023
'@LastModified October 15, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.stack?view=netframework-4.8.1

'@Remarks
'Singleton class for Stack containing factory methods and static members.

Option Explicit

Private Type TStackSingleton
   StackSingleton As DotNetLib.StackSingleton
End Type

Private this As TStackSingleton

Private Sub Class_Initialize()
   Set this.StackSingleton = New DotNetLib.StackSingleton
End Sub

Private Sub Class_Terminate()
   Set this.StackSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the Stack class that is empty and has the specified initial capacity or the default initial capacity, whichever is greater.")
Public Function Create(Optional ByVal initialCapacity As Long = 10) As DotNetLib.Stack
Attribute Create.VB_Description = "Initializes a new instance of the Stack class that is empty and has the specified initial capacity or the default initial capacity, whichever is greater."
   Set Create = this.StackSingleton.Create(initialCapacity)
End Function

'@Description("Initializes a new instance of the Stack class that contains elements copied from the specified collection and has the same initial capacity as the number of elements copied.")
Public Function Create2(ByRef col As DotNetLib.ICollection) As DotNetLib.Stack
Attribute Create2.VB_Description = "Initializes a new instance of the Stack class that contains elements copied from the specified collection and has the same initial capacity as the number of elements copied."
   Set Create2 = this.StackSingleton.Create2(col)
End Function

'@Description("Returns a synchronized (thread safe) wrapper for the Stack.")
Public Function Synchronized(ByRef col As DotNetLib.Stack) As DotNetLib.Stack
Attribute Synchronized.VB_Description = "Returns a synchronized (thread safe) wrapper for the Stack."
   Set Synchronized = this.StackSingleton.Synchronized(col)
End Function
