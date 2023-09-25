VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Strings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents text as a sequence of UTF-16 code units."
'@Folder("VBADotNetLib.System")
'@PredeclaredId
'@ModuleDescription("Represents text as a sequence of UTF-16 code units.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 23, 2023
'@LastModified September 24, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

'@Remarks
' Singleton class for String static members and factory methods

Option Explicit

Private Type TString
    String As DotNetLib.StringSingleton
End Type

Private this As TString

Private Sub Class_Initialize()
    Set this.String = New DotNetLib.StringSingleton
End Sub

Private Sub Class_Terminate()
    Set this.String = Nothing
End Sub

'@Static
'@Description("Represents the empty string. This field is read-only.")
'@Remarks Rename from Empty due to VBA reserved word
Public Property Get EmptyString() As String
Attribute EmptyString.VB_Description = "Represents the empty string. This field is read-only."
    EmptyString = this.String.EmptyString
End Property

'@Static
'@Description("Compares two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare(ByVal strA As String, ByVal strB As String, Optional ByVal ignoreCase As Boolean = False) As Long
Attribute Compare.VB_Description = "Compares two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order."
   Compare = this.String.Compare(strA, strB, ignoreCase)
End Function

'@Static
'@Description("Compares two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare2(ByVal strA As String, ByVal strB As String, ByVal comparisonType As StringComparison) As Long
Attribute Compare2.VB_Description = "Compares two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order."
    Compare2 = this.String.Compare2(strA, strB, comparisonType)
End Function

'@Static
'@Description("Compares two specified String objects, ignoring or honoring their case, and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare3(ByVal strA As String, ByVal strB As String, ByVal ignoreCase As Boolean, ByVal culture As DotNetLib.CultureInfo) As Long
Attribute Compare3.VB_Description = "Compares two specified String objects, ignoring or honoring their case, and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order."
    Compare3 = this.String.Compare3(strA, strB, ignoreCase, culture)
End Function

'@Static
'@Description("Compares two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two strings to each other in the sort order.")
Public Function Compare4(ByVal strA As String, ByVal strB As String, ByVal culture As DotNetLib.CultureInfo, ByVal options As CompareOptions) As Long
Attribute Compare4.VB_Description = "Compares two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two strings to each other in the sort order."
    Compare4 = this.String.Compare4(strA, strB, culture, options)
End Function

'@Static
'@Description("Compares substrings of two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare5(ByVal strA As String, ByVal indexA As Long, ByVal strB As String, ByVal indexB As Long, ByVal length As Long, Optional ByVal ignoreCase As Boolean = False) As Long
Attribute Compare5.VB_Description = "Compares substrings of two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order."
    Compare5 = this.String.Compare5(strA, indexA, strB, indexB, length, ignoreCase)
End Function

'@Static
'@Description("Compares substrings of two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare6(ByVal strA As String, ByVal indexA As Long, ByVal strB As String, ByVal indexB As Long, ByVal length As Long, ByVal comparisonType As StringComparison) As Long
Attribute Compare6.VB_Description = "Compares substrings of two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order."
    Compare6 = this.String.Compare6(strA, indexA, strB, indexB, length, comparisonType)
End Function

'@Static
'@Description("Compares substrings of two specified String objects, ignoring or honoring their case and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare7(ByVal strA As String, ByVal indexA As Long, ByVal strB As String, ByVal indexB As Long, ByVal length As Long, ByVal ignoreCase As Boolean, ByVal culture As DotNetLib.CultureInfo) As Long
Attribute Compare7.VB_Description = "Compares substrings of two specified String objects, ignoring or honoring their case and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order."
    Compare7 = this.String.Compare7(strA, indexA, strB, indexB, length, ignoreCase, culture)
End Function

'@Static
'@Description("Compares substrings of two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two substrings to each other in the sort order.")
Public Function Compare8(ByVal strA As String, ByVal indexA As Long, ByVal strB As String, ByVal indexB As Long, ByVal length As Long, ByVal culture As DotNetLib.CultureInfo, ByVal options As CompareOptions) As Long
Attribute Compare8.VB_Description = "Compares substrings of two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two substrings to each other in the sort order."
    Compare8 = this.String.Compare8(strA, indexA, strB, indexB, length, culture, options)
End Function

'@Static
'@Description("Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string.")
Public Function CompareOrdinal(ByVal strA As String, ByVal strB As String) As Long
Attribute CompareOrdinal.VB_Description = "Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string."
    CompareOrdinal = this.String.CompareOrdinal(strA, strB)
End Function

'@Static
'@Description("Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring.")
Public Function CompareOrdinal2(ByVal strA As String, ByVal indexA As Long, ByVal strB As String, ByVal indexB As Long, ByVal length As Long) As Long
Attribute CompareOrdinal2.VB_Description = "Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring."
    CompareOrdinal2 = this.String.CompareOrdinal2(strA, indexA, strB, indexB, length)
End Function

'@Static
'@Description("Creates a new instance of String with the same value as a specified String.")
Public Function Copy(ByVal str As String) As String
Attribute Copy.VB_Description = "Creates a new instance of String with the same value as a specified String."
   Copy = this.String.Copy(str)
End Function

'@Static
'@Description("Determines whether two specified String objects have the same value.")
Public Function Equals(ByVal a As String, ByVal b As String) As Boolean
Attribute Equals.VB_Description = "Determines whether two specified String objects have the same value."
   Equals = this.String.Equals(a, b)
End Function

'@Static
'@Description("Determines whether two specified String objects have the same value. A parameter specifies the culture, case, and sort rules used in the comparison.")
Public Function Equals2(ByVal a As String, ByVal b As String, ByVal comparisonType As StringComparison) As Boolean
Attribute Equals2.VB_Description = "Determines whether two specified String objects have the same value. A parameter specifies the culture, case, and sort rules used in the comparison."
   Equals2 = this.String.Equals2(a, b, comparisonType)
End Function

'@Static
'@Description("Replaces the format item in a specified string with the string representation of a corresponding object in a specified array.")
Public Function Format(ByVal pFormat As String, ParamArray args() As Variant) As String
Attribute Format.VB_Description = "Replaces the format item in a specified string with the string representation of a corresponding object in a specified array."
    Dim params() As Variant
    params = args
    Format = this.String.Format(pFormat, params)
End Function

'@Static
'@Description("Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information.")
Public Function Format2(ByVal provider As IFormatProvider, ByVal pFormat As String, ParamArray args() As Variant) As String
Attribute Format2.VB_Description = "Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information."
    Dim params() As Variant
    params = args
    Format2 = this.String.Format2(provider, pFormat, params)
End Function

'@Static
'@Description("Indicates whether the specified string is null or an empty string ("").")
Public Function IsNullOrEmpty(ByVal value As String) As Boolean
Attribute IsNullOrEmpty.VB_Description = "Indicates whether the specified string is null or an empty string ("")."
   IsNullOrEmpty = this.String.IsNullOrEmpty(value)
End Function

'@Static
'@Description("Indicates whether a specified string is null, empty, or consists only of white-space characters.")
Public Function IsNullOrWhiteSpace(ByVal value As String) As Boolean
Attribute IsNullOrWhiteSpace.VB_Description = "Indicates whether a specified string is null, empty, or consists only of white-space characters."
   IsNullOrWhiteSpace = this.String.IsNullOrWhiteSpace(value)
End Function
