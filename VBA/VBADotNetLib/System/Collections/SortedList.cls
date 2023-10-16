VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a collection of key/value pairs that are sorted by the keys and are accessible by key and by index."
'@Folder "VBADotNetLib.System.Collections"
'@PredeclaredId
'@ModuleDescription("Represents a collection of key/value pairs that are sorted by the keys and are accessible by key and by index.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 16, 2023
'@LastModified October 16, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist?view=netframework-4.8.1

'@Remarks
'Singleton class for SortedList containing factory methods and static members.

Option Explicit

Private Type TSortedListSingleton
   SortedListSingleton As DotNetLib.SortedListSingleton
End Type

Private this As TSortedListSingleton

Private Sub Class_Initialize()
   Set this.SortedListSingleton = New DotNetLib.SortedListSingleton
End Sub

Private Sub Class_Terminate()
   Set this.SortedListSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the SortedList class that is empty, has the specified initial capacity, and is sorted according to the IComparable interface implemented by each key added to the SortedList object.")
Public Function Create(Optional ByVal initialCapacity As Long = 0) As DotNetLib.SortedList
   Set Create = this.SortedListSingleton.Create(initialCapacity)
End Function

'@Description("Initializes a new instance of the SortedList class that is empty, has the specified initial capacity, and is sorted according to the specified IComparer interface.")
Public Function Create2(ByVal comparer As mscorlib.IComparer, Optional ByVal Capacity As Long = 0) As DotNetLib.SortedList
   Set Create2 = this.SortedListSingleton.Create2(comparer, Capacity)
End Function

'@Description("Initializes a new instance of the SortedList class that contains elements copied from the specified dictionary, has the same initial capacity as the number of elements copied, and is sorted according to the specified IComparer interface.")
Public Function Create3(ByVal d As DotNetLib.IDictionary, Optional ByRef comparer As mscorlib.IComparer) As DotNetLib.SortedList
   Set Create3 = this.SortedListSingleton.Create3(d, comparer)
End Function

