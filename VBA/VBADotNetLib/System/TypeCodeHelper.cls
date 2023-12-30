VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypeCodeHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@ModuleDescription("TypeCode helper class to convert a TypeCode to string value.")
'@Folder "VBADotNetLib.System"
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 30, 2023
'@LastModified December 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.typecode?view=netframework-4.8.1

Option Explicit

'@DefaultMember
'@Description("Converts the value of a DateTimeKind to string value.")
Public Property Get ToString(ByVal value As mscorlib.TypeCode) As String
   Select Case value
      Case TypeCode_Boolean: ToString = "Boolean"
      Case TypeCode_Byte: ToString = "Byte"
      Case TypeCode_Char: ToString = "Char"
      Case TypeCode_DateTime: ToString = "DateTime"
      Case TypeCode_DBNull: ToString = "DBNull"
      Case TypeCode_Decimal: ToString = "Decimal"
      Case TypeCode_Double: ToString = "Double"
      Case TypeCode_Empty: ToString = "Empty"
      Case TypeCode_Int16: ToString = "Int16"
      Case TypeCode_Int32: ToString = "Int32"
      Case TypeCode_Int64: ToString = "Int64"
      Case TypeCode_Object: ToString = "Object"
      Case TypeCode_SByte: ToString = "SByte"
      Case TypeCode_Single: ToString = "Single"
      Case TypeCode_String: ToString = "String"
      Case TypeCode_UInt16: ToString = "UInt16"
      Case TypeCode_UInt32: ToString = "UInt32"
      Case TypeCode_UInt64: ToString = "UInt64"
   End Select
End Property
