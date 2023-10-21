VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionaryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Defines a dictionary key/value pair that can be set or retrieved."
'@Folder "VBADotNetLib.System.Collections"
'@PredeclaredId
'@ModuleDescription("Defines a dictionary key/value pair that can be set or retrieved.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 19, 2023
'@LastModified October 20, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.dictionaryentry?view=netframework-4.8.1

'@Remarks
'Singleton class for DictionaryEntry containing factory methods and static members.

Option Explicit

Const DictionaryEntryType As String = "DotNetLib.System.Collections.DictionaryEntry"

Private Type TDictionaryEntrySingleton
    DictionaryEntrySingleton As DotNetLib.DictionaryEntrySingleton
    DictionaryEntryType As DotNetLib.Type
End Type

Private this As TDictionaryEntrySingleton

Private Sub Class_Initialize()
   Set this.DictionaryEntrySingleton = New DotNetLib.DictionaryEntrySingleton
    Set this.DictionaryEntryType = Types.GetType(DictionaryEntryType)
End Sub

Private Sub Class_Terminate()
   Set this.DictionaryEntrySingleton = Nothing
End Sub

'@Description("Initializes an instance of the DictionaryEntry type with the specified key and value.")
Public Function Create(ByVal pDictionaryEntry As Variant) As DotNetLib.DictionaryEntry
Attribute Create.VB_Description = "Initializes an instance of the DictionaryEntry type with the specified key and value."
   Set Create = this.DictionaryEntrySingleton.Create(pDictionaryEntry)
End Function

'@Description("Initializes an instance of the DictionaryEntry type from a mscorlib.DictionaryEntry containing the specified key and value.")
Public Function Create2(ByRef pDictionaryEntry As mscorlib.DictionaryEntry) As DotNetLib.DictionaryEntry
Attribute Create2.VB_Description = "Initializes an instance of the DictionaryEntry type from a mscorlib.DictionaryEntry containing the specified key and value."
   Set Create2 = this.DictionaryEntrySingleton.Create2(pDictionaryEntry)
End Function

'@Description("Initializes an instance of the DictionaryEntry type with the specified key and value.")
Public Function Create3(ByVal Key As Variant, ByVal Value As Variant) As DotNetLib.DictionaryEntry
Attribute Create3.VB_Description = "Initializes an instance of the DictionaryEntry type with the specified key and value."
   Set Create3 = this.DictionaryEntrySingleton.Create3(Key, Value)
End Function

'--------------------------------------------------------------------
'Extension
'Added to get the type for arrays etc.
'--------------------------------------------------------------------
'@Static
'@DefaultMember
'@Description("Gets the String Type.")
Public Function GetType() As DotNetLib.Type
Attribute GetType.VB_Description = "Gets the String Type."
Attribute GetType.VB_UserMemId = 0
    Set GetType = this.DictionaryEntryType
End Function
