VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBAString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents text as a sequence of UTF-16 code units."
'@Folder "VBADotNetLib.System"
'@PredeclaredId
'@ModuleDescription("Represents text as a sequence of UTF-16 code units.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 23, 2023
'@LastModified November 14, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

'@Remarks
' Singleton class for String static members and factory methods

Option Explicit

Const StringType As String = "System.String"

Private Type TString
    BStringSingleton As DotNetLib.BStringSingleton
'    StringSingleton As DotNetLib.StringSingleton
    StringType   As DotNetLib.Type
End Type

Private this As TString

Private Sub Class_Initialize()
    Set this.BStringSingleton = New DotNetLib.BStringSingleton
'    Set this.StringSingleton = New DotNetLib.StringSingleton
    Set this.StringType = Types.GetType(StringType)
End Sub

Private Sub Class_Terminate()
    Set this.BStringSingleton = Nothing
End Sub

''@Description("Initializes a new instance of the String class to the Unicode characters indicated in the specified string.")
'Public Function Create(ByVal value As String) As DotNetLib.String
'   Set Create = this.StringSingleton.Create(value)
'End Function
'
''@Description("Initializes a new instance of the String class to the value indicated by a specified Unicode character repeated a specified number of times.")
'Public Function Create2(ByVal character As String, ByVal Count As Long) As DotNetLib.String
'   Set Create2 = this.StringSingleton.Create_2(character, Count)
'End Function
'
''@Description("Initializes a new instance of the String class to the value indicated by an string of Unicode characters, a starting character position within that array, and a length")
'Public Function Create3(ByVal value As String, ByVal startIndex As Long, ByVal length As Long) As DotNetLib.String
'   Set Create3 = this.StringSingleton.Create_3(value, startIndex, length)
'End Function
'
''@Description("Initializes a new instance of the String class to the Unicode characters indicated in the specified string.")
'Public Function Create4(ByRef value As DotNetLib.String) As DotNetLib.String
'   Set Create4 = this.StringSingleton.Create_4(value)
'End Function
'
''@Description("Initializes a new instance of the String class to the value indicated by a specified Unicode character repeated a specified number of times.")
'Public Function Create5(ByRef character As DotNetLib.String, ByVal Count As Long) As DotNetLib.String
'   Set Create5 = this.BStringSingleton.Create_5(character, Count)
'End Function
'
''@Description("Initializes a new instance of the String class to the value indicated by an string of Unicode characters, a starting character position within that array, and a length")
'Public Function Create6(ByRef value As DotNetLib.String, ByVal startIndex As Long, ByVal length As Long) As DotNetLib.String
'   Set Create6 = this.StringSingleton.Create_6(value, startIndex, length)
'End Function

'@Static
'@Description("Represents the empty string. This field is read-only.")
'@Remarks Rename from Empty due to VBA reserved word
Public Property Get EmptyString() As String
Attribute EmptyString.VB_Description = "Represents the empty string. This field is read-only."
    EmptyString = this.BStringSingleton.EmptyString
End Property

'@Static
'@Description("Compares two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare(ByVal strA As String, ByVal strB As String, Optional ByVal ignoreCase As Boolean = False) As Long
Attribute Compare.VB_Description = "Compares two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order."
   Compare = this.BStringSingleton.Compare(strA, strB, ignoreCase)
End Function

'@Static
'@Description("Compares two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare2(ByVal strA As String, ByVal strB As String, ByVal comparisonType As StringComparison) As Long
Attribute Compare2.VB_Description = "Compares two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order."
    Compare2 = this.BStringSingleton.Compare2(strA, strB, comparisonType)
End Function

'@Static
'@Description("Compares two specified String objects, ignoring or honoring their case, and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare3(ByVal strA As String, ByVal strB As String, ByVal ignoreCase As Boolean, ByVal culture As DotNetLib.CultureInfo) As Long
Attribute Compare3.VB_Description = "Compares two specified String objects, ignoring or honoring their case, and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order."
    Compare3 = this.BStringSingleton.Compare3(strA, strB, ignoreCase, culture)
End Function

'@Static
'@Description("Compares two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two strings to each other in the sort order.")
Public Function Compare4(ByVal strA As String, ByVal strB As String, ByVal culture As DotNetLib.CultureInfo, ByVal options As CompareOptions) As Long
Attribute Compare4.VB_Description = "Compares two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two strings to each other in the sort order."
    Compare4 = this.BStringSingleton.Compare4(strA, strB, culture, options)
End Function

'@Static
'@Description("Compares substrings of two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare5(ByVal strA As String, ByVal indexA As Long, ByVal strB As String, ByVal indexB As Long, ByVal length As Long, Optional ByVal ignoreCase As Boolean = False) As Long
Attribute Compare5.VB_Description = "Compares substrings of two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order."
    Compare5 = this.BStringSingleton.Compare5(strA, indexA, strB, indexB, length, ignoreCase)
End Function

'@Static
'@Description("Compares substrings of two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare6(ByVal strA As String, ByVal indexA As Long, ByVal strB As String, ByVal indexB As Long, ByVal length As Long, ByVal comparisonType As StringComparison) As Long
Attribute Compare6.VB_Description = "Compares substrings of two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order."
    Compare6 = this.BStringSingleton.Compare6(strA, indexA, strB, indexB, length, comparisonType)
End Function

'@Static
'@Description("Compares substrings of two specified String objects, ignoring or honoring their case and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare7(ByVal strA As String, ByVal indexA As Long, ByVal strB As String, ByVal indexB As Long, ByVal length As Long, ByVal ignoreCase As Boolean, ByVal culture As DotNetLib.CultureInfo) As Long
Attribute Compare7.VB_Description = "Compares substrings of two specified String objects, ignoring or honoring their case and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order."
    Compare7 = this.BStringSingleton.Compare7(strA, indexA, strB, indexB, length, ignoreCase, culture)
End Function

'@Static
'@Description("Compares substrings of two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two substrings to each other in the sort order.")
Public Function Compare8(ByVal strA As String, ByVal indexA As Long, ByVal strB As String, ByVal indexB As Long, ByVal length As Long, ByVal culture As DotNetLib.CultureInfo, ByVal options As CompareOptions) As Long
Attribute Compare8.VB_Description = "Compares substrings of two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two substrings to each other in the sort order."
    Compare8 = this.BStringSingleton.Compare8(strA, indexA, strB, indexB, length, culture, options)
End Function

'@Static
'@Description("Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string.")
Public Function CompareOrdinal(ByVal strA As String, ByVal strB As String) As Long
Attribute CompareOrdinal.VB_Description = "Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string."
    CompareOrdinal = this.BStringSingleton.CompareOrdinal(strA, strB)
End Function

'@Static
'@Description("Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring.")
Public Function CompareOrdinal2(ByVal strA As String, ByVal indexA As Long, ByVal strB As String, ByVal indexB As Long, ByVal length As Long) As Long
Attribute CompareOrdinal2.VB_Description = "Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring."
    CompareOrdinal2 = this.BStringSingleton.CompareOrdinal2(strA, indexA, strB, indexB, length)
End Function

'@Static
'@Description("Creates a new instance of String with the same value as a specified String.")
Public Function Copy(ByVal str As String) As String
Attribute Copy.VB_Description = "Creates a new instance of String with the same value as a specified String."
   Copy = this.BStringSingleton.Copy(str)
End Function

'@Static
'@Description("Determines whether two specified String objects have the same value.")
Public Function Equals(ByVal a As String, ByVal b As String) As Boolean
Attribute Equals.VB_Description = "Determines whether two specified String objects have the same value."
   Equals = this.BStringSingleton.Equals(a, b)
End Function

'@Static
'@Description("Determines whether two specified String objects have the same value. A parameter specifies the culture, case, and sort rules used in the comparison.")
Public Function Equals2(ByVal a As String, ByVal b As String, ByVal comparisonType As StringComparison) As Boolean
Attribute Equals2.VB_Description = "Determines whether two specified String objects have the same value. A parameter specifies the culture, case, and sort rules used in the comparison."
   Equals2 = this.BStringSingleton.Equals2(a, b, comparisonType)
End Function

'@Static
'@Description("Replaces the format item in a specified string with the string representation of a corresponding object in a specified array.")
Public Function Format(ByVal pFormat As String, ParamArray args() As Variant) As String
Attribute Format.VB_Description = "Replaces the format item in a specified string with the string representation of a corresponding object in a specified array."
    Dim params() As Variant
    params = args
    Format = this.BStringSingleton.Format(pFormat, params)
End Function

'@Static
'@Description("Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information.")
Public Function Format2(ByVal provider As IFormatProvider, ByVal pFormat As String, ParamArray args() As Variant) As String
Attribute Format2.VB_Description = "Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information."
    Dim params() As Variant
    params = args
    Format2 = this.BStringSingleton.Format2(provider, pFormat, params)
End Function

'@Static
'@Description("Indicates whether the specified string is null or an empty string ("").")
Public Function IsNullOrEmpty(ByVal value As String) As Boolean
Attribute IsNullOrEmpty.VB_Description = "Indicates whether the specified string is null or an empty string ("")."
   IsNullOrEmpty = this.BStringSingleton.IsNullOrEmpty(value)
End Function

'@Static
'@Description("Indicates whether a specified string is null, empty, or consists only of white-space characters.")
Public Function IsNullOrWhiteSpace(ByVal value As String) As Boolean
Attribute IsNullOrWhiteSpace.VB_Description = "Indicates whether a specified string is null, empty, or consists only of white-space characters."
   IsNullOrWhiteSpace = this.BStringSingleton.IsNullOrWhiteSpace(value)
End Function

'@Static
'@Description("Concatenates all the elements of a string array, using the specified separator between each element.")
Public Function Join(ByVal separator As String, ByRef value() As String) As String
Attribute Join.VB_Description = "Concatenates all the elements of a string array, using the specified separator between each element."
   Join = this.BStringSingleton.Join(separator, value)
End Function

'@Static
'@Description("Concatenates the elements of an object array, using the specified separator between each element.")
Public Function Join2(ByVal separator As String, ByRef values() As Variant) As String
Attribute Join2.VB_Description = "Concatenates the elements of an object array, using the specified separator between each element."
   Join2 = this.BStringSingleton.Join2(separator, values)
End Function

'@Static
'@Description("Concatenates the members of a constructed IEnumerable<T> collection of type String, using the specified separator between each member.")
Public Function Join3(ByVal separator As String, ByVal stringValues As mscorlib.IEnumerable) As String
Attribute Join3.VB_Description = "Concatenates the members of a constructed IEnumerable<T> collection of type String, using the specified separator between each member."
   Join3 = this.BStringSingleton.Join3(separator, stringValues)
End Function

'@Static
'@Description("Concatenates the specified elements of a string array, using the specified separator between each element.")
Public Function Join4(ByVal separator As String, ByRef value() As String, ByVal startIndex As Long, ByVal count As Long) As String
Attribute Join4.VB_Description = "Concatenates the specified elements of a string array, using the specified separator between each element."
   Join4 = this.BStringSingleton.Join4(separator, value, startIndex, count)
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
    Set GetType = this.StringType
End Function