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
'@LastModified October 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.dictionaryentry?view=netframework-4.8.1

'@Remarks
'Singleton class for DictionaryEntry containing factory methods and static members.

'@Todo Remove Create factory methods as no longer required, rename Assign to Create?

Option Explicit

'For creating an Array of mscorlib.DictionaryEntry
Const DictionaryEntryType As String = "System.Collections.DictionaryEntry"

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

'@Todo : Rename to Create
'@Static
'@Description("Initializes an instance of the mscorlib.DictionaryEntry type with the specified key and value.")
Public Sub Assign(ByRef pDictionaryEntry As mscorlib.DictionaryEntry, ByVal pKey As Variant, ByVal pValue As Variant)
Attribute Assign.VB_Description = "Initializes an instance of the mscorlib.DictionaryEntry type with the specified key and value."
    Call this.DictionaryEntrySingleton.Assign(pDictionaryEntry, pKey, pValue)
End Sub

'@Static
'@Description("Gets or sets the key in the key/value pair.")
Public Property Get Key(ByRef pDictionaryEntry As mscorlib.DictionaryEntry) As Variant
Attribute Key.VB_Description = "Gets or sets the key in the key/value pair."
    Call AssignVariable(Key, this.DictionaryEntrySingleton.GetKey(pDictionaryEntry))
End Property

'@Static
'@Description("Gets or sets the key in the key/value pair.")
Public Property Let Key(ByRef pDictionaryEntry As mscorlib.DictionaryEntry, ByVal pKey As Variant)
Attribute Key.VB_Description = "Gets or sets the key in the key/value pair."
    Call this.DictionaryEntrySingleton.SetKey(pDictionaryEntry, pKey)
End Property

'@Static
'@Description("Gets or sets the key in the key/value pair.")
Public Property Set Key(ByRef pDictionaryEntry As mscorlib.DictionaryEntry, ByVal pKey As Variant)
Attribute Key.VB_Description = "Gets or sets the key in the key/value pair."
    Call this.DictionaryEntrySingleton.SetKey(pDictionaryEntry, pKey)
End Property

'@Static
'@Description("Gets or sets the value in the key/value pair.")
Public Property Get Value(ByRef pDictionaryEntry As mscorlib.DictionaryEntry) As Variant
Attribute Value.VB_Description = "Gets or sets the value in the key/value pair."
    Call AssignVariable(Value, this.DictionaryEntrySingleton.GetValue(pDictionaryEntry))
End Property

'@Description("Gets or sets the value in the key/value pair.")
Public Property Let Value(ByRef pDictionaryEntry As mscorlib.DictionaryEntry, ByVal pValue As Variant)
Attribute Value.VB_Description = "Gets or sets the value in the key/value pair."
    Call this.DictionaryEntrySingleton.SetValue(pDictionaryEntry, pValue)
End Property

'@Description("Gets or sets the value in the key/value pair.")
Public Property Set Value(ByRef pDictionaryEntry As mscorlib.DictionaryEntry, ByVal pValue As Variant)
Attribute Value.VB_Description = "Gets or sets the value in the key/value pair."
    Call this.DictionaryEntrySingleton.SetValue(pDictionaryEntry, pValue)
End Property

'--------------------------------------------------------------------
'Extension
'Added to get the type for arrays etc.
'--------------------------------------------------------------------
'@Static
'@DefaultMember
'@Description("Gets the mscorlib.DictionaryEntry Type.")
Public Function GetType() As DotNetLib.Type
Attribute GetType.VB_Description = "Gets the mscorlib.DictionaryEntry Type."
Attribute GetType.VB_UserMemId = 0
    Set GetType = this.DictionaryEntryType
End Function

' @Reference
' https://stackoverflow.com/questions/51080637/write-result-of-function-to-a-variable-where-result-could-be-an-object/53925677#53925677
' https://stackoverflow.com/a/53925677/10759363
' Assign x to y regardless of object or primitive
Private Sub AssignVariable(ByRef x As Variant, ByRef y As Variant)
    If IsObject(y) Then
        Set x = y
    Else
        x = y
    End If
End Sub

'@Todo remove Create factory methods and  rename Assign to Create

'@Static
'@Description("Initializes an instance of the DictionaryEntry type with the specified key and value.")
Public Function Create(ByVal pDictionaryEntry As Variant) As DotNetLib.DictionaryEntry
Attribute Create.VB_Description = "Initializes an instance of the DictionaryEntry type with the specified key and value."
   Set Create = this.DictionaryEntrySingleton.Create(pDictionaryEntry)
End Function

'Static
'@Description("Initializes an instance of the DictionaryEntry type from a mscorlib.DictionaryEntry containing the specified key and value.")
Public Function Create2(ByRef pDictionaryEntry As mscorlib.DictionaryEntry) As DotNetLib.DictionaryEntry
Attribute Create2.VB_Description = "Initializes an instance of the DictionaryEntry type from a mscorlib.DictionaryEntry containing the specified key and value."
   Set Create2 = this.DictionaryEntrySingleton.Create2(pDictionaryEntry)
End Function

'@Static
'@Description("Initializes an instance of the DictionaryEntry type with the specified key and value.")
Public Function Create3(ByVal pKey As Variant, ByVal pValue As Variant) As DotNetLib.DictionaryEntry
Attribute Create3.VB_Description = "Initializes an instance of the DictionaryEntry type with the specified key and value."
   Set Create3 = this.DictionaryEntrySingleton.Create3(pKey, pValue)
End Function

''Static
''@Description("Gets the key in the key/value pair.")
'Public Function GetKey(ByRef pDictionaryEntry As mscorlib.DictionaryEntry) As Variant
'    Call AssignVariable(GetKey, this.DictionaryEntrySingleton.GetKey(pDictionaryEntry))
'End Function

''Static
''@Description("Gets the value in the key/value pair.")
'Public Function GetValue(ByRef pDictionaryEntry As mscorlib.DictionaryEntry) As Variant
'    Call AssignVariable(GetValue, this.DictionaryEntrySingleton.GetValue(pDictionaryEntry))
'End Function

''@Description("Sets the key in the key/value pair.")
'Public Sub SetKey(ByRef pDictionaryEntry As mscorlib.DictionaryEntry, ByVal pKey As Variant)
'    Call this.DictionaryEntrySingleton.SetKey(pDictionaryEntry, pKey)
'End Sub

''@Description("Sets the value in the key/value pair.")
'Public Sub SetValue(ByRef pDictionaryEntry As mscorlib.DictionaryEntry, ByVal pValue As Variant)
'    Call this.DictionaryEntrySingleton.SetValue(pDictionaryEntry, pValue)
'End Sub
