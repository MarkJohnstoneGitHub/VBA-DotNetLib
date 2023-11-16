VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Objects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Supports all classes in the .NET class hierarchy and provides low-level services to derived classes. This is the ultimate base class of all .NET classes; it is the root of the type hierarchy."
'@Folder "VBADotNetLib.System"
'@ModuleDescription("Supports all classes in the .NET class hierarchy and provides low-level services to derived classes. This is the ultimate base class of all .NET classes; it is the root of the type hierarchy.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 8, 2023
'@LastModified October 10, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.object?view=netframework-4.8.1

'@Remarks
' Singleton class for Regex containing constructors and static members.

Option Explicit

Private Type TObject
   ObjectSingleton As DotNetLib.ObjectSingleton
End Type

Private this As TObject

Private Sub Class_Initialize()
   Set this.ObjectSingleton = New DotNetLib.ObjectSingleton
End Sub

Private Sub Class_Terminate()
   Set this.ObjectSingleton = Nothing
End Sub

Public Function Create(Optional ByVal obj As Variant = Nothing) As DotNetLib.Object
    Set Create = this.ObjectSingleton.Create(obj)
End Function

Public Function Equals(ByVal objA As Variant, ByVal objB As Variant) As Boolean
    Equals = this.ObjectSingleton.Equals(objA, objB)
End Function
