VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Implements the IList interface using an array whose size is dynamically increased as required."
'@Folder "VBADotNetLib.System.Collections"
'@PredeclaredId
'@ModuleDescription("Implements the IList interface using an array whose size is dynamically increased as required.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 5, 2023
'@LastModified October 6, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8.1

'@Remarks
'Singleton class for ArrayList containing factory methods and static members.

Option Explicit

Private Type TArrayListSingleton
   ArrayListSingleton As DotNetLib.ArrayListSingleton
End Type

Private this As TArrayListSingleton

Private Sub Class_Initialize()
   Set this.ArrayListSingleton = New DotNetLib.ArrayListSingleton
End Sub

Private Sub Class_Terminate()
   Set this.ArrayListSingleton = Nothing
End Sub

'@Static
'@Description("Initializes a new instance of the ArrayList class that is empty and has the default or specified initial capacity.")
Public Function Create(Optional ByVal Capacity As Long = 0) As DotNetLib.ArrayList
Attribute Create.VB_Description = "Initializes a new instance of the ArrayList class that is empty and has the default or specified initial capacity."
   Set Create = this.ArrayListSingleton.Create(Capacity)
End Function

'@Static
'@Description("Initializes a new instance of the ArrayList class that contains elements copied from the specified collection and that has the same initial capacity as the number of elements copied.")
Public Function Create2(ByVal c As mscorlib.ICollection) As DotNetLib.ArrayList
Attribute Create2.VB_Description = "Initializes a new instance of the ArrayList class that contains elements copied from the specified collection and that has the same initial capacity as the number of elements copied."
   Set Create2 = this.ArrayListSingleton.Create2(c)
End Function

'@Static
'@Description("Creates an ArrayList wrapper for a specific IList.")
Public Function Adapter(ByVal pList As mscorlib.IList) As DotNetLib.ArrayList
Attribute Adapter.VB_Description = "Creates an ArrayList wrapper for a specific IList."
   Set Adapter = this.ArrayListSingleton.Adapter(pList)
End Function

'@Static
'@Description("Returns an ArrayList wrapper with a fixed size.")
Public Function FixedSize(ByVal pList As DotNetLib.ArrayList) As DotNetLib.ArrayList
Attribute FixedSize.VB_Description = "Returns an ArrayList wrapper with a fixed size."
   Set FixedSize = this.ArrayListSingleton.FixedSize(pList)
End Function

'@Static
'@Description("Returns an IList wrapper with a fixed size.")
Public Function FixedSize2(ByRef List As mscorlib.IList) As IList
Attribute FixedSize2.VB_Description = "Returns an IList wrapper with a fixed size."
   Set FixedSize2 = this.ArrayListSingleton.FixedSize2(List)
End Function

'@Static
'@Description("Returns a read-only ArrayList wrapper.")
Public Function ReadOnly(ByVal pList As DotNetLib.ArrayList) As DotNetLib.ArrayList
Attribute ReadOnly.VB_Description = "Returns a read-only ArrayList wrapper."
   Set ReadOnly = this.ArrayListSingleton.ReadOnly(pList)
End Function

'@Static
'@Description("Returns a read-only IList wrapper.")
Public Function ReadOnly2(ByVal pList As mscorlib.IList) As mscorlib.IList
Attribute ReadOnly2.VB_Description = "Returns a read-only IList wrapper."
   Set ReadOnly2 = this.ArrayListSingleton.ReadOnly2(pList)
End Function

'@Static
'@Description("Returns an ArrayList whose elements are copies of the specified value.")
Public Function Repeat(ByVal value As Variant, ByVal Count As Long) As DotNetLib.ArrayList
Attribute Repeat.VB_Description = "Returns an ArrayList whose elements are copies of the specified value."
   Set Repeat = this.ArrayListSingleton.Repeat(value, Count)
End Function

'@Static
'@Description("Returns an ArrayList wrapper that is synchronized (thread safe).")
Public Function Synchronized(ByVal pList As DotNetLib.ArrayList) As DotNetLib.ArrayList
Attribute Synchronized.VB_Description = "Returns an ArrayList wrapper that is synchronized (thread safe)."
   Set Synchronized = this.ArrayListSingleton.Synchronized(pList)
End Function

'@Static
'@Description("Returns an IList wrapper that is synchronized (thread safe).")
Public Function Synchronized2(ByVal pList As mscorlib.IList) As mscorlib.IList
Attribute Synchronized2.VB_Description = "Returns an IList wrapper that is synchronized (thread safe)."
   Set Synchronized2 = this.ArrayListSingleton.Synchronized2(pList)
End Function
