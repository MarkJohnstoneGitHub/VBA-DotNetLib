VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a mutable string of characters. This class cannot be inherited."
'@Folder "VBADotNetLib.System.Text"
'@PredeclaredId
'@ModuleDescription("Represents a mutable string of characters. This class cannot be inherited.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 30, 2023
'@LastModified October 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder?view=netframework-4.8.1

Option Explicit

Private Type TStringBuilderSingleton
   StringBuilderSingleton As DotNetLib.StringBuilderSingleton
End Type

Private this As TStringBuilderSingleton

Private Sub Class_Initialize()
   Set this.StringBuilderSingleton = New DotNetLib.StringBuilderSingleton
End Sub

Private Sub Class_Terminate()
   Set this.StringBuilderSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the StringBuilder class using the specified string and capacity.")
Public Function Create(Optional ByVal value As String = VBA.vbNullString, Optional ByVal Capacity As Long = 16) As DotNetLib.StringBuilder
Attribute Create.VB_Description = "Initializes a new instance of the StringBuilder class using the specified string and capacity."
   Set Create = this.StringBuilderSingleton.Create(value, Capacity)
End Function

'@Description("Initializes a new instance of the StringBuilder class that starts with a specified capacity and can grow to a specified maximum.")
Public Function Create2(ByVal Capacity As Long, ByVal MaxCapacity As Long) As DotNetLib.StringBuilder
Attribute Create2.VB_Description = "Initializes a new instance of the StringBuilder class that starts with a specified capacity and can grow to a specified maximum."
   Set Create2 = this.StringBuilderSingleton.Create2(Capacity, MaxCapacity)
End Function

'@Description("Initializes a new instance of the StringBuilder class from the specified substring and capacity.")
Public Function Create3(ByVal value As String, ByVal startIndex As Long, ByVal Length As Long, ByVal Capacity As Long) As DotNetLib.StringBuilder
Attribute Create3.VB_Description = "Initializes a new instance of the StringBuilder class from the specified substring and capacity."
   Set Create3 = this.StringBuilderSingleton.Create3(value, startIndex, Length, Capacity)
End Function
