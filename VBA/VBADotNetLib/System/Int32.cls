VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Int32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a 32-bit signed integer."
'@Folder "VBADotNetLib.System"
'@ModuleDescription("Represents a 32-bit signed integer.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 12, 2023
'@LastModified February 14, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.int32?view=netframework-4.8.1

'@Remarks
' Singleton class for Int32 containing constructors and static members.

Option Explicit

Const Int32Type As String = "System.Int32"

Private Type TInt32
    Int32Type       As DotNetLib.Type
    Int32Singleton  As DotNetLib.Int32Singleton
End Type

Private this As TInt32

Private Sub Class_Initialize()
    Set this.Int32Singleton = New DotNetLib.Int32Singleton
    Set this.Int32Type = Types.GetType(Int32Type)
End Sub

Private Sub Class_Terminate()
    Set this.Int32Type = Nothing
End Sub

'@Description("Represents the largest possible value of an Int32. This field is constant.")
Public Property Get MaxValue() As Long
Attribute MaxValue.VB_Description = "Represents the largest possible value of an Int32. This field is constant."
   MaxValue = this.Int32Singleton.MaxValue()
End Property

'@Description("Represents the smallest possible value of Int32. This field is constant.")
Public Property Get MinValue() As Long
Attribute MinValue.VB_Description = "Represents the smallest possible value of Int32. This field is constant."
   MinValue = this.Int32Singleton.MinValue()
End Property

'@Description("Converts the string representation of a number to its 32-bit signed integer equivalent. A return value indicates whether the conversion succeeded.")
Public Function TryParse(ByVal s As String, ByRef result As Long) As Boolean
Attribute TryParse.VB_Description = "Converts the string representation of a number to its 32-bit signed integer equivalent. A return value indicates whether the conversion succeeded."
   TryParse = this.Int32Singleton.TryParse(s, result)
End Function

'@Description("Converts the string representation of a number in a specified style and culture-specific format to its 32-bit signed integer equivalent. A return value indicates whether the conversion succeeded.")
Public Function TryParse2(ByRef s As String, ByVal style As mscorlib.NumberStyles, ByVal provider As mscorlib.IFormatProvider, ByRef outResult As Long) As Boolean
Attribute TryParse2.VB_Description = "Converts the string representation of a number in a specified style and culture-specific format to its 32-bit signed integer equivalent. A return value indicates whether the conversion succeeded."
   TryParse2 = this.Int32Singleton.TryParse_2(s, style, provider, outResult)
End Function

'--------------------------------------------------------------------
'Extension
'Added to get the type for arrays etc.
'--------------------------------------------------------------------
'@Static
'@DefaultMember
'@Description("Gets the Int32 Type.")
Public Function GetType() As DotNetLib.Type
Attribute GetType.VB_Description = "Gets the Int32 Type."
Attribute GetType.VB_UserMemId = 0
    Set GetType = this.Int32Type
End Function
