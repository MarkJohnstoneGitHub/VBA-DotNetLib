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
'@Folder "VBADotNetLib.System"
'@ModuleDescription("Represents text as a sequence of UTF-16 code units.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 11, 2023
'@LastModified November 16, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

'@Remarks
' Singleton class for String static members and factory methods

Option Explicit

Private Type TStringSingleton
   StringSingleton As DotNetLib.StringSingleton
End Type

Private this As TStringSingleton

Private Sub Class_Initialize()
   Set this.StringSingleton = New DotNetLib.StringSingleton
End Sub

Private Sub Class_Terminate()
   Set this.StringSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the String class to the Unicode characters indicated in the specified string.")
Public Function Create(ByVal value As String) As DotNetLib.String
Attribute Create.VB_Description = "Initializes a new instance of the String class to the Unicode characters indicated in the specified string."
   Set Create = this.StringSingleton.Create(value)
End Function

'@Description("Initializes a new instance of the String class to the value indicated by a specified Unicode character repeated a specified number of times.")
Public Function Create2(ByVal character As String, ByVal count As Long) As DotNetLib.String
Attribute Create2.VB_Description = "Initializes a new instance of the String class to the value indicated by a specified Unicode character repeated a specified number of times."
   Set Create2 = this.StringSingleton.Create_2(character, count)
End Function

'@Description("Initializes a new instance of the String class to the value indicated by an string of Unicode characters, a starting character position within that array, and a length")
Public Function Create3(ByVal value As String, ByVal startIndex As Long, ByVal length As Long) As DotNetLib.String
Attribute Create3.VB_Description = "Initializes a new instance of the String class to the value indicated by an string of Unicode characters, a starting character position within that array, and a length"
   Set Create3 = this.StringSingleton.Create_3(value, startIndex, length)
End Function

'@Description("Initializes a new instance of the String class to the Unicode characters indicated in the specified string.")
Public Function Create4(ByVal value As DotNetLib.String) As DotNetLib.String
Attribute Create4.VB_Description = "Initializes a new instance of the String class to the Unicode characters indicated in the specified string."
   Set Create4 = this.StringSingleton.Create_4(value)
End Function

'@Description("Initializes a new instance of the String class to the value indicated by a specified Unicode character repeated a specified number of times.")
Public Function Create5(ByVal character As DotNetLib.String, ByVal count As Long) As DotNetLib.String
Attribute Create5.VB_Description = "Initializes a new instance of the String class to the value indicated by a specified Unicode character repeated a specified number of times."
   Set Create5 = this.StringSingleton.Create_5(character, count)
End Function

'@Description("Initializes a new instance of the String class to the value indicated by an string of Unicode characters, a starting character position within that array, and a length")
Public Function Create6(ByVal value As DotNetLib.String, ByVal startIndex As Long, ByVal length As Long) As DotNetLib.String
Attribute Create6.VB_Description = "Initializes a new instance of the String class to the value indicated by an string of Unicode characters, a starting character position within that array, and a length"
   Set Create6 = this.StringSingleton.Create_6(value, startIndex, length)
End Function

'@Description("Initializes a new instance of the String class to the value indicated by an string of Unicode characters, converting any escaped characters in the input string.")
Public Function CreateUnescape(ByVal value As String) As DotNetLib.String
Attribute CreateUnescape.VB_Description = "Initializes a new instance of the String class to the value indicated by an string of Unicode characters, converting any escaped characters in the input string."
   Set CreateUnescape = this.StringSingleton.CreateUnescape(value)
End Function

'@Description("Represents the empty string. This field is read-only.")
Public Property Get EmptyString() As DotNetLib.String
Attribute EmptyString.VB_Description = "Represents the empty string. This field is read-only."
   Set EmptyString = this.StringSingleton.EmptyString()
End Property

'@Description("Compares two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare(ByVal strA As DotNetLib.String, ByVal strB As DotNetLib.String, Optional ByVal ignoreCase As Boolean = False) As Long
Attribute Compare.VB_Description = "Compares two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order."
   Compare = this.StringSingleton.Compare(strA, strB, ignoreCase)
End Function

'@Description("Compares two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare2(ByVal strA As DotNetLib.String, ByVal strB As DotNetLib.String, ByVal comparisonType As mscorlib.StringComparison) As Long
Attribute Compare2.VB_Description = "Compares two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order."
   Compare2 = this.StringSingleton.Compare2(strA, strB, comparisonType)
End Function

'@Description("Compares two specified String objects, ignoring or honoring their case, and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare3(ByVal strA As DotNetLib.String, ByVal strB As DotNetLib.String, ByVal ignoreCase As Boolean, ByVal culture As DotNetLib.CultureInfo) As Long
Attribute Compare3.VB_Description = "Compares two specified String objects, ignoring or honoring their case, and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order."
   Compare3 = this.StringSingleton.Compare3(strA, strB, ignoreCase, culture)
End Function

'@Description("Compares two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two strings to each other in the sort order.")
Public Function Compare4(ByVal strA As DotNetLib.String, ByVal strB As DotNetLib.String, ByVal culture As DotNetLib.CultureInfo, ByVal options As mscorlib.CompareOptions) As Long
Attribute Compare4.VB_Description = "Compares two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two strings to each other in the sort order."
   Compare4 = this.StringSingleton.Compare4(strA, strB, culture, options)
End Function

'@Description("Compares substrings of two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare5(ByVal strA As DotNetLib.String, ByVal indexA As Long, ByVal strB As DotNetLib.String, ByVal indexB As Long, ByVal length As Long, Optional ByVal ignoreCase As Boolean = False) As Long
Attribute Compare5.VB_Description = "Compares substrings of two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order."
   Compare5 = this.StringSingleton.Compare5(strA, indexA, strB, indexB, length, ignoreCase)
End Function

'@Description("Compares substrings of two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare6(ByVal strA As DotNetLib.String, ByVal indexA As Long, ByVal strB As DotNetLib.String, ByVal indexB As Long, ByVal length As Long, ByVal comparisonType As mscorlib.StringComparison) As Long
Attribute Compare6.VB_Description = "Compares substrings of two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order."
   Compare6 = this.StringSingleton.Compare6(strA, indexA, strB, indexB, length, comparisonType)
End Function

'@Description("Compares substrings of two specified String objects, ignoring or honoring their case and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare7(ByVal strA As DotNetLib.String, ByVal indexA As Long, ByVal strB As DotNetLib.String, ByVal indexB As Long, ByVal length As Long, ByVal ignoreCase As Boolean, ByVal culture As DotNetLib.CultureInfo) As Long
Attribute Compare7.VB_Description = "Compares substrings of two specified String objects, ignoring or honoring their case and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order."
   Compare7 = this.StringSingleton.Compare7(strA, indexA, strB, indexB, length, ignoreCase, culture)
End Function

'@Description("Compares substrings of two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two substrings to each other in the sort order.")
Public Function Compare8(ByVal strA As DotNetLib.String, ByVal indexA As Long, ByVal strB As DotNetLib.String, ByVal indexB As Long, ByVal length As Long, ByVal culture As DotNetLib.CultureInfo, ByVal options As mscorlib.CompareOptions) As Long
Attribute Compare8.VB_Description = "Compares substrings of two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two substrings to each other in the sort order."
   Compare8 = this.StringSingleton.Compare8(strA, indexA, strB, indexB, length, culture, options)
End Function

'@Description("Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string.")
Public Function CompareOrdinal(ByVal strA As DotNetLib.String, ByVal strB As DotNetLib.String) As Long
Attribute CompareOrdinal.VB_Description = "Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string."
   CompareOrdinal = this.StringSingleton.CompareOrdinal(strA, strB)
End Function

'@Description("Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring.")
Public Function CompareOrdinal2(ByVal strA As DotNetLib.String, ByVal indexA As Long, ByVal strB As DotNetLib.String, ByVal indexB As Long, ByVal length As Long) As Long
Attribute CompareOrdinal2.VB_Description = "Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring."
   CompareOrdinal2 = this.StringSingleton.CompareOrdinal2(strA, indexA, strB, indexB, length)
End Function

'@Description("Creates the string representation of a specified object.")
Public Function Concat(ByVal arg0 As Variant) As DotNetLib.String
Attribute Concat.VB_Description = "Creates the string representation of a specified object."
   Set Concat = this.StringSingleton.Concat(arg0)
End Function

'@Description("Concatenates the string representations of two specified objects.")
Public Function Concat2(ByVal arg0 As Variant, ByVal arg1 As Variant) As DotNetLib.String
Attribute Concat2.VB_Description = "Concatenates the string representations of two specified objects."
   Set Concat2 = this.StringSingleton.Concat2(arg0, arg1)
End Function

'@Description("Concatenates the string representations of three specified objects.")
Public Function Concat3(ByVal arg0 As Variant, ByVal arg1 As Variant, ByVal arg2 As Variant) As DotNetLib.String
Attribute Concat3.VB_Description = "Concatenates the string representations of three specified objects."
   Set Concat3 = this.StringSingleton.Concat3(arg0, arg1, arg2)
End Function

'@Description("Concatenates the string representations of the elements in a specified Object array.")
Public Function Concat4(ByRef values() As Variant) As DotNetLib.String
Attribute Concat4.VB_Description = "Concatenates the string representations of the elements in a specified Object array."
   Set Concat4 = this.StringSingleton.Concat4(values)
End Function

'@Description("Concatenates the members of a constructed IEnumerable<T> collection of type String.")
Public Function Concat5(ByVal stringValues As IEnumerable) As DotNetLib.String
Attribute Concat5.VB_Description = "Concatenates the members of a constructed IEnumerable<T> collection of type String."
   Set Concat5 = this.StringSingleton.Concat5(stringValues)
End Function

'@Description("Concatenates the members of an IEnumerable<T> implementation.")
Public Function Concat6(ByVal values As IEnumerable) As DotNetLib.String
Attribute Concat6.VB_Description = "Concatenates the members of an IEnumerable<T> implementation."
   Set Concat6 = this.StringSingleton.Concat6(values)
End Function

'@Description("Concatenates two specified instances of String.")
Public Function Concat7(ByVal str0 As String, ByVal str1 As String) As DotNetLib.String
Attribute Concat7.VB_Description = "Concatenates two specified instances of String."
   Set Concat7 = this.StringSingleton.Concat7(str0, str1)
End Function

'@Description("Concatenates three specified instances of String.")
Public Function Concat8(ByVal str0 As String, ByVal str1 As String, ByVal str2 As String) As DotNetLib.String
Attribute Concat8.VB_Description = "Concatenates three specified instances of String."
   Set Concat8 = this.StringSingleton.Concat8(str0, str1, str2)
End Function

'@Description("Concatenates four specified instances of String.")
Public Function Concat9(ByVal str0 As String, ByVal str1 As String, ByVal str2 As String, ByVal str3 As String) As DotNetLib.String
Attribute Concat9.VB_Description = "Concatenates four specified instances of String."
   Set Concat9 = this.StringSingleton.Concat9(str0, str1, str2, str3)
End Function

'@Description("Concatenates the elements of a specified String array.")
Public Function Concat10(ByRef values() As String) As DotNetLib.String
Attribute Concat10.VB_Description = "Concatenates the elements of a specified String array."
   Set Concat10 = this.StringSingleton.Concat10(values)
End Function

'@Description("Concatenates two specified instances of String.")
Public Function Concat12(ByVal str0 As DotNetLib.String, ByVal str1 As DotNetLib.String) As DotNetLib.String
Attribute Concat12.VB_Description = "Concatenates two specified instances of String."
   Set Concat12 = this.StringSingleton.Concat12(str0, str1)
End Function

'@Description("Concatenates three specified instances of String.")
Public Function Concat13(ByVal str0 As DotNetLib.String, ByVal str1 As DotNetLib.String, ByVal str2 As DotNetLib.String) As DotNetLib.String
Attribute Concat13.VB_Description = "Concatenates three specified instances of String."
   Set Concat13 = this.StringSingleton.Concat13(str0, str1, str2)
End Function

'@Description("Concatenates four specified instances of String.")
Public Function Concat14(ByVal str0 As DotNetLib.String, ByVal str1 As DotNetLib.String, ByVal str2 As DotNetLib.String, ByVal str3 As DotNetLib.String) As DotNetLib.String
Attribute Concat14.VB_Description = "Concatenates four specified instances of String."
   Set Concat14 = this.StringSingleton.Concat14(str0, str1, str2, str3)
End Function

'@Description("Creates a new instance of String with the same value as a specified String.")
Public Function Copy(ByVal str As String) As DotNetLib.String
Attribute Copy.VB_Description = "Creates a new instance of String with the same value as a specified String."
   Set Copy = this.StringSingleton.Copy(str)
End Function

'@Description("Creates a new instance of String with the same value as a specified String.")
Public Function Copy2(ByVal str As DotNetLib.String) As DotNetLib.String
Attribute Copy2.VB_Description = "Creates a new instance of String with the same value as a specified String."
   Set Copy2 = this.StringSingleton.Copy_2(str)
End Function

'@Description("Determines whether two specified String objects have the same value.")
Public Function Equals(ByVal a As DotNetLib.String, ByVal b As DotNetLib.String) As Boolean
Attribute Equals.VB_Description = "Determines whether two specified String objects have the same value."
   Equals = this.StringSingleton.Equals(a, b)
End Function

'@Description("Determines whether two specified String objects have the same value. A parameter specifies the culture, case, and sort rules used in the comparison.")
Public Function Equals2(ByVal a As DotNetLib.String, ByVal b As DotNetLib.String, ByVal comparisonType As mscorlib.StringComparison) As Boolean
Attribute Equals2.VB_Description = "Determines whether two specified String objects have the same value. A parameter specifies the culture, case, and sort rules used in the comparison."
   Equals2 = this.StringSingleton.Equals2(a, b, comparisonType)
End Function

'@Description("Replaces the format item in a specified string with the string representation of a corresponding object in a specified array.")
Public Function Format(ByVal pFormat As String, ByRef args() As Variant) As DotNetLib.String
Attribute Format.VB_Description = "Replaces the format item in a specified string with the string representation of a corresponding object in a specified array."
   Set Format = this.StringSingleton.Format(pFormat, args)
End Function

'@Description("Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information.")
Public Function Format2(ByVal provider As IFormatProvider, ByVal pFormat As String, ByRef args() As Variant) As DotNetLib.String
Attribute Format2.VB_Description = "Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information."
   Set Format2 = this.StringSingleton.Format2(provider, pFormat, args)
End Function

'@Description("Replaces the format item in a specified string with the string representation of a corresponding object in a specified array.")
Public Function Format3(ByVal pFormat As DotNetLib.String, ByRef args() As Variant) As DotNetLib.String
Attribute Format3.VB_Description = "Replaces the format item in a specified string with the string representation of a corresponding object in a specified array."
   Set Format3 = this.StringSingleton.Format3(pFormat, args)
End Function

'@Description("Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information.")
Public Function Format4(ByVal provider As IFormatProvider, ByVal pFormat As DotNetLib.String, ByRef args() As Variant) As DotNetLib.String
Attribute Format4.VB_Description = "Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information."
   Set Format4 = this.StringSingleton.Format4(provider, pFormat, args)
End Function

'@Description("Indicates whether the specified string is null or an empty string ("").")
Public Function IsNullOrEmpty(ByVal value As DotNetLib.String) As Boolean
Attribute IsNullOrEmpty.VB_Description = "Indicates whether the specified string is null or an empty string ("")."
   IsNullOrEmpty = this.StringSingleton.IsNullOrEmpty(value)
End Function

'@Description("Indicates whether a specified string is null, empty, or consists only of white-space characters.")
Public Function IsNullOrWhiteSpace(ByVal value As DotNetLib.String) As Boolean
Attribute IsNullOrWhiteSpace.VB_Description = "Indicates whether a specified string is null, empty, or consists only of white-space characters."
   IsNullOrWhiteSpace = this.StringSingleton.IsNullOrWhiteSpace(value)
End Function

'@Description("Concatenates all the elements of a string array, using the specified separator between each element.")
Public Function Join(ByVal separator As String, ByRef value() As String) As DotNetLib.String
Attribute Join.VB_Description = "Concatenates all the elements of a string array, using the specified separator between each element."
   Set Join = this.StringSingleton.Join(separator, value)
End Function

'@Description("Concatenates the elements of an object array, using the specified separator between each element.")
Public Function Join2(ByVal separator As String, ByRef values() As Variant) As DotNetLib.String
Attribute Join2.VB_Description = "Concatenates the elements of an object array, using the specified separator between each element."
   Set Join2 = this.StringSingleton.Join_2(separator, values)
End Function

'@Description("Concatenates the members of a constructed IEnumerable<T> collection of type String, using the specified separator between each member.")
Public Function Join3(ByVal separator As String, ByRef stringValues As IEnumerable) As DotNetLib.String
Attribute Join3.VB_Description = "Concatenates the members of a constructed IEnumerable<T> collection of type String, using the specified separator between each member."
   Set Join3 = this.StringSingleton.Join3(separator, stringValues)
End Function

'@Description("Concatenates the specified elements of a string array, using the specified separator between each element.")
Public Function Join4(ByVal separator As String, ByRef value() As String, ByVal startIndex As Long, ByVal pCount As Long) As DotNetLib.String
Attribute Join4.VB_Description = "Concatenates the specified elements of a string array, using the specified separator between each element."
   Set Join4 = this.StringSingleton.Join4(separator, value, startIndex, pCount)
End Function
