VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a strongly typed list of objects that can be accessed by index. Provides methods to search, sort, and manipulate lists."
'@Folder "VBADotNetLib.System.Collections"
'@ModuleDescription("Represents a strongly typed list of objects that can be accessed by index. Provides methods to search, sort, and manipulate lists.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 1, 2023
'@LastModified October 3, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1?view=netframework-4.8.1

'@Remarks
'Singleton class for List containing factory methods and static members.

Option Explicit

Private Type TListSingleton
   ListSingleton As DotNetLib.ListSingleton
End Type

Private this As TListSingleton

Private Sub Class_Initialize()
   Set this.ListSingleton = New DotNetLib.ListSingleton
End Sub

Private Sub Class_Terminate()
   Set this.ListSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the List<object> class that is empty and has the default initial capacity.")
Public Function Create() As DotNetLib.List
Attribute Create.VB_Description = "Initializes a new instance of the List<object> class that is empty and has the default initial capacity."
   Set Create = this.ListSingleton.Create()
End Function

'@Description("Initializes a new instance of the List class of the type provided that is empty and has the default or specified initial capacity.")
Public Function Create2(ByVal listType As Variant, Optional ByVal pCapacity As Long = 0) As DotNetLib.List
Attribute Create2.VB_Description = "Initializes a new instance of the List class of the type provided that is empty and has the default or specified initial capacity."
   Set Create2 = this.ListSingleton.Create2(listType, pCapacity)
End Function
