VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a strongly typed list of objects that can be accessed by index. Provides methods to search, sort, and manipulate lists."
'@Folder "VBADotNetLib.System.Collections"
'@ModuleDescription("Represents a strongly typed list of strings that can be accessed by index. Provides methods to search, sort, and manipulate lists.")
'@PredeclaredId
'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 3, 2023
'@LastModified October 3, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1?view=netframework-4.8.1

'@Remarks
'Singleton class for List<String> containing factory methods and static members.

Option Explicit

Private Type TListStringSingleton
   ListStringSingleton As DotNetLib.ListStringSingleton
End Type

Private this As TListStringSingleton

Private Sub Class_Initialize()
   Set this.ListStringSingleton = New DotNetLib.ListStringSingleton
End Sub

Private Sub Class_Terminate()
   Set this.ListStringSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the List<String> class that is empty and has the default or specified initial capacity.")
Public Function Create(Optional ByVal pCapacity As Long = 0) As DotNetLib.ListString
Attribute Create.VB_Description = "Initializes a new instance of the List<String> class that is empty and has the default or specified initial capacity."
   Set Create = this.ListStringSingleton.Create2(pCapacity)
End Function


''@Description("Initializes a new instance of the List<String> class that is empty and has the default or specified initial capacity.")
'Public Function Create2(ByVal pCapacity As Long) As DotNetLib.ListString
'   Set Create2 = this.ListStringSingleton.Create2(pCapacity)
'End Function
