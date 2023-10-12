VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Types"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents type declarations: class types, interface types, array types, value types, enumeration types, type parameters, generic type definitions, and open or closed constructed generic types."
'@ModuleDescription("Represents type declarations: class types, interface types, array types, value types, enumeration types, type parameters, generic type definitions, and open or closed constructed generic types.")
'@Folder "VBADotNetLib.System"
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 8, 2023
'@LastModified October 10, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.type?view=netframework-4.8.1

'@Remarks
' Singleton class for Type containing constructors and static members.

Option Explicit

Private Type TTypeSingleton
    TypeSingleton As DotNetLib.TypeSingleton
End Type

Private this As TTypeSingleton

Private Sub Class_Initialize()
    Set this.TypeSingleton = New DotNetLib.TypeSingleton
End Sub

Private Sub Class_Terminate()
    Set this.TypeSingleton = Nothing
End Sub

'Bug Issue throwing exception when throwOnError is false https://github.com/dotnet/runtime/issues/12376 or behaving as intended?
'//   throwOnError:
'//     true to throw an exception if the type cannot be found; false to return null.
'//     Specifying false also suppresses some other exception conditions, but not all
'//     of them. See the Exceptions section.
'@Description("Gets the Type with the specified name, specifying whether to throw an exception if the type is not found and whether to perform a case-sensitive search.")
Public Function GetType(ByVal typeName As String, Optional ByVal throwOnError As Boolean = False, Optional ByVal ignoreCase As Boolean = False) As DotNetLib.Type
Attribute GetType.VB_Description = "Gets the Type with the specified name, specifying whether to throw an exception if the type is not found and whether to perform a case-sensitive search."
    Set GetType = this.TypeSingleton.GetType_3(typeName, throwOnError, ignoreCase)
End Function

'@Description("Gets the underlying type code of the specified Type.")
Public Function GetTypeCode(ByVal pType As DotNetLib.Type) As TypeCode
Attribute GetTypeCode.VB_Description = "Gets the underlying type code of the specified Type."
   GetTypeCode = this.TypeSingleton.GetTypeCode(pType)
End Function


''@Description("Gets the Type with the specified name, performing a case-sensitive search.")
'Public Function GetType(ByVal typeName As String) As DotNetLib.Type
'   Set GetType = this.TypeSingleton.GetType(typeName)
'End Function
'
''@Description("Gets the Type with the specified name, performing a case-sensitive search and specifying whether to throw an exception if the type is not found.")
'Public Function GetType2(ByVal typeName As String, ByVal throwOnError As Boolean) As DotNetLib.Type
'   Set GetType2 = this.TypeSingleton.GetType_2(typeName, throwOnError)
'End Function
'
''@Description("Gets the Type with the specified name, specifying whether to throw an exception if the type is not found and whether to perform a case-sensitive search.")
'Public Function GetType3(ByVal typeName As String, ByVal throwOnError As Boolean, ByVal ignoreCase As Boolean) As DotNetLib.Type
'   Set GetType3 = this.TypeSingleton.GetType_3(typeName, throwOnError, ignoreCase)
'End Function
