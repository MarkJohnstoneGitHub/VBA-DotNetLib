VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Char"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a character as a UTF-16 code unit."
'@Folder("VBADotNetLib.System")
'@ModuleDescription("Represents a character as a UTF-16 code unit.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 15, 2023
'@LastModified October 20, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.char?view=netframework-4.8.1

'@Remarks
' Singleton class for Char containing factory methods and static members.

Option Explicit

Const CharType As String = "System.Char"

Private Type TChar
    CharType As DotNetLib.Type
    CharSingleton As DotNetLib.CharSingleton
End Type

Private this As TChar

Private Sub Class_Initialize()
    Set this.CharType = Types.GetType(CharType)
    Set this.CharSingleton = New DotNetLib.CharSingleton
End Sub

Private Sub Class_Terminate()
    Set this.CharType = Nothing
    Set this.CharSingleton = Nothing
End Sub

'@Static
'@Description("Converts the specified Unicode code point into a UTF-16 encoded string.")
Public Function ConvertFromUtf32(ByVal utf32 As Long) As String
Attribute ConvertFromUtf32.VB_Description = "Converts the specified Unicode code point into a UTF-16 encoded string."
   ConvertFromUtf32 = this.CharSingleton.ConvertFromUtf32(utf32)
End Function

'@Static
'@Description("Converts the value of a UTF-16 encoded character or surrogate pair at a specified position in a string into a Unicode code point.")
Public Function ConvertToUtf32(ByRef s As String, ByVal index As Long) As Long
Attribute ConvertToUtf32.VB_Description = "Converts the value of a UTF-16 encoded character or surrogate pair at a specified position in a string into a Unicode code point."
   ConvertToUtf32 = this.CharSingleton.ConvertToUtf32(s, index)
End Function

'@Static
'@Description("Converts the numeric Unicode character at the specified position in a specified string to a double-precision floating point number.")
Public Function GetNumericValue(ByRef s As String, ByVal index As Long) As Double
Attribute GetNumericValue.VB_Description = "Converts the numeric Unicode character at the specified position in a specified string to a double-precision floating point number."
   GetNumericValue = this.CharSingleton.GetNumericValue(s, index)
End Function

'@Static
'@Description("Categorizes a Unicode character into a group identified by one of the UnicodeCategory values.")
Public Function GetUnicodeCategory(ByRef s As String, ByVal index As Long) As UnicodeCategory
Attribute GetUnicodeCategory.VB_Description = "Categorizes a Unicode character into a group identified by one of the UnicodeCategory values."
   GetUnicodeCategory = this.CharSingleton.GetUnicodeCategory(s, index)
End Function

'@Static
'@Description("Indicates whether the character at the specified position in a specified string is categorized as a control character.")
Public Function IsControl(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsControl.VB_Description = "Indicates whether the character at the specified position in a specified string is categorized as a control character."
   IsControl = this.CharSingleton.IsControl(s, index)
End Function

'@Static
'@Description("Indicates whether the character at the specified position in a specified string is categorized as a decimal digit.")
Public Function IsDigit(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsDigit.VB_Description = "Indicates whether the character at the specified position in a specified string is categorized as a decimal digit."
   IsDigit = this.CharSingleton.IsDigit(s, index)
End Function

'@Static
'@Description("Indicates whether the Char object at the specified position in a string is a high surrogate.")
Public Function IsHighSurrogate(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsHighSurrogate.VB_Description = "Indicates whether the Char object at the specified position in a string is a high surrogate."
   IsHighSurrogate = this.CharSingleton.IsHighSurrogate(s, index)
End Function

'@Static
'@Description("Indicates whether the character at the specified position in a specified string is categorized as a Unicode letter.")
Public Function IsLetter(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsLetter.VB_Description = "Indicates whether the character at the specified position in a specified string is categorized as a Unicode letter."
   IsLetter = this.CharSingleton.IsLetter(s, index)
End Function

'@Static
'@Description("Indicates whether the character at the specified position in a specified string is categorized as a letter or a decimal digit.")
Public Function IsLetterOrDigit(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsLetterOrDigit.VB_Description = "Indicates whether the character at the specified position in a specified string is categorized as a letter or a decimal digit."
   IsLetterOrDigit = this.CharSingleton.IsLetterOrDigit(s, index)
End Function

'@Static
'@Description("Indicates whether the character at the specified position in a specified string is categorized as a lowercase letter.")
Public Function IsLower(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsLower.VB_Description = "Indicates whether the character at the specified position in a specified string is categorized as a lowercase letter."
   IsLower = this.CharSingleton.IsLower(s, index)
End Function

'@Static
'@Description("Indicates whether the Char object at the specified position in a string is a low surrogate.")
Public Function IsLowSurrogate(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsLowSurrogate.VB_Description = "Indicates whether the Char object at the specified position in a string is a low surrogate."
   IsLowSurrogate = this.CharSingleton.IsLowSurrogate(s, index)
End Function

'@Static
'@Description("Indicates whether the character at the specified position in a specified string is categorized as a number.")
Public Function IsNumber(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsNumber.VB_Description = "Indicates whether the character at the specified position in a specified string is categorized as a number."
   IsNumber = this.CharSingleton.IsNumber(s, index)
End Function

'@Static
'@Description("Indicates whether the character at the specified position in a specified string is categorized as a punctuation mark.")
Public Function IsPunctuation(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsPunctuation.VB_Description = "Indicates whether the character at the specified position in a specified string is categorized as a punctuation mark."
   IsPunctuation = this.CharSingleton.IsPunctuation(s, index)
End Function

'@Static
'@Description("Indicates whether the character at the specified position in a specified string is categorized as a separator character.")
Public Function IsSeparator(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsSeparator.VB_Description = "Indicates whether the character at the specified position in a specified string is categorized as a separator character."
   IsSeparator = this.CharSingleton.IsSeparator(s, index)
End Function

'@Static
'@Description("Indicates whether the character at the specified position in a specified string has a surrogate code unit.")
Public Function IsSurrogate(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsSurrogate.VB_Description = "Indicates whether the character at the specified position in a specified string has a surrogate code unit."
   IsSurrogate = this.CharSingleton.IsSurrogate(s, index)
End Function

'@Static
'@Description("Indicates whether two adjacent Char objects at a specified position in a string form a surrogate pair.")
Public Function IsSurrogatePair(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsSurrogatePair.VB_Description = "Indicates whether two adjacent Char objects at a specified position in a string form a surrogate pair."
   IsSurrogatePair = this.CharSingleton.IsSurrogatePair(s, index)
End Function

'@Static
'@Description("Indicates whether the character at the specified position in a specified string is categorized as a symbol character.")
Public Function IsSymbol(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsSymbol.VB_Description = "Indicates whether the character at the specified position in a specified string is categorized as a symbol character."
   IsSymbol = this.CharSingleton.IsSymbol(s, index)
End Function

'@Static
'@Description("Indicates whether the character at the specified position in a specified string is categorized as an uppercase letter.")
Public Function IsUpper(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsUpper.VB_Description = "Indicates whether the character at the specified position in a specified string is categorized as an uppercase letter."
   IsUpper = this.CharSingleton.IsUpper(s, index)
End Function

'@Static
'@Description("Indicates whether the character at the specified position in a specified string is categorized as white space.")
Public Function IsWhiteSpace(ByRef s As String, ByVal index As Long) As Boolean
Attribute IsWhiteSpace.VB_Description = "Indicates whether the character at the specified position in a specified string is categorized as white space."
   IsWhiteSpace = this.CharSingleton.IsWhiteSpace(s, index)
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
    Set GetType = this.CharType
End Function
