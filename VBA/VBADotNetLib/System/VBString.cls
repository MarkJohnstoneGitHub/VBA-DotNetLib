VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents text as a sequence of UTF-16 code units."
'@IgnoreModule ParameterCanBeByVal
'@Folder "VBADotNetLib.System"
'@PredeclaredId
'@ModuleDescription("Represents text as a sequence of UTF-16 code units.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 23, 2023
'@LastModified January 23, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

'@Remarks
' Singleton class for String static members and factory methods

Option Explicit

Const StringType As String = "System.String"

Private Type TString
    BStringSingleton As DotNetLib.BStringSingleton
    StringType   As DotNetLib.Type
End Type

Private this As TString

Private Sub Class_Initialize()
    Set this.BStringSingleton = New DotNetLib.BStringSingleton
    Set this.StringType = Types.GetType(StringType)
End Sub

Private Sub Class_Terminate()
    Set this.BStringSingleton = Nothing
End Sub

'@Static
'@Description("Represents the empty string. This field is read-only.")
'@Remarks Rename from Empty due to VBA reserved word
Public Property Get EmptyString() As String
Attribute EmptyString.VB_Description = "Represents the empty string. This field is read-only."
    EmptyString = this.BStringSingleton.EmptyString
End Property

'@Static
'@Description("Compares two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare(ByRef strA As String, ByRef strB As String, Optional ByVal ignoreCase As Boolean = False) As Long
Attribute Compare.VB_Description = "Compares two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order."
   Compare = this.BStringSingleton.Compare(strA, strB, ignoreCase)
End Function

'@Static
'@Description("Compares two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare2(ByRef strA As String, ByRef strB As String, ByVal comparisonType As mscorlib.StringComparison) As Long
Attribute Compare2.VB_Description = "Compares two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order."
    Compare2 = this.BStringSingleton.Compare2(strA, strB, comparisonType)
End Function

'@Static
'@Description("Compares two specified String objects, ignoring or honoring their case, and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare3(ByRef strA As String, ByRef strB As String, ByVal ignoreCase As Boolean, ByVal culture As DotNetLib.CultureInfo) As Long
Attribute Compare3.VB_Description = "Compares two specified String objects, ignoring or honoring their case, and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order."
    Compare3 = this.BStringSingleton.Compare3(strA, strB, ignoreCase, culture)
End Function

'@Static
'@Description("Compares two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two strings to each other in the sort order.")
Public Function Compare4(ByRef strA As String, ByRef strB As String, ByVal culture As DotNetLib.CultureInfo, ByVal options As mscorlib.CompareOptions) As Long
Attribute Compare4.VB_Description = "Compares two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two strings to each other in the sort order."
    Compare4 = this.BStringSingleton.Compare4(strA, strB, culture, options)
End Function

'@Static
'@Description("Compares substrings of two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare5(ByRef strA As String, ByVal indexA As Long, ByRef strB As String, ByVal indexB As Long, ByVal pLength As Long, Optional ByVal ignoreCase As Boolean = False) As Long
Attribute Compare5.VB_Description = "Compares substrings of two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order."
    Compare5 = this.BStringSingleton.Compare5(strA, indexA, strB, indexB, pLength, ignoreCase)
End Function

'@Static
'@Description("Compares substrings of two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare6(ByRef strA As String, ByVal indexA As Long, ByRef strB As String, ByVal indexB As Long, ByVal pLength As Long, ByVal comparisonType As mscorlib.StringComparison) As Long
Attribute Compare6.VB_Description = "Compares substrings of two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order."
    Compare6 = this.BStringSingleton.Compare6(strA, indexA, strB, indexB, pLength, comparisonType)
End Function

'@Static
'@Description("Compares substrings of two specified String objects, ignoring or honoring their case and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order.")
Public Function Compare7(ByRef strA As String, ByVal indexA As Long, ByRef strB As String, ByVal indexB As Long, ByVal pLength As Long, ByVal ignoreCase As Boolean, ByVal culture As DotNetLib.CultureInfo) As Long
Attribute Compare7.VB_Description = "Compares substrings of two specified String objects, ignoring or honoring their case and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order."
    Compare7 = this.BStringSingleton.Compare7(strA, indexA, strB, indexB, pLength, ignoreCase, culture)
End Function

'@Static
'@Description("Compares substrings of two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two substrings to each other in the sort order.")
Public Function Compare8(ByRef strA As String, ByVal indexA As Long, ByRef strB As String, ByVal indexB As Long, ByVal pLength As Long, ByVal culture As DotNetLib.CultureInfo, ByVal options As mscorlib.CompareOptions) As Long
Attribute Compare8.VB_Description = "Compares substrings of two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two substrings to each other in the sort order."
    Compare8 = this.BStringSingleton.Compare8(strA, indexA, strB, indexB, pLength, culture, options)
End Function

'@Static
'@Description("Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string.")
Public Function CompareOrdinal(ByRef strA As String, ByRef strB As String) As Long
Attribute CompareOrdinal.VB_Description = "Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string."
    CompareOrdinal = this.BStringSingleton.CompareOrdinal(strA, strB)
End Function

'@Static
'@Description("Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring.")
Public Function CompareOrdinal2(ByRef strA As String, ByVal indexA As Long, ByRef strB As String, ByVal indexB As Long, ByVal pLength As Long) As Long
Attribute CompareOrdinal2.VB_Description = "Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring."
    CompareOrdinal2 = this.BStringSingleton.CompareOrdinal2(strA, indexA, strB, indexB, pLength)
End Function

'@Static
'@Description("Concatenates the elements of a specified String array.")
Public Function Concat(ByRef values() As String) As String
Attribute Concat.VB_Description = "Concatenates the elements of a specified String array."
   Concat = this.BStringSingleton.Concat(values)
End Function

'@Static
'@Description("Concatenates two specified instances of String.")
Public Function Concat2(ByRef str0 As String, ByRef str1 As String) As String
Attribute Concat2.VB_Description = "Concatenates two specified instances of String."
   Concat2 = this.BStringSingleton.Concat2(str0, str1)
End Function

'@Static
'@Description("Concatenates three specified instances of String.")
Public Function Concat3(ByRef str0 As String, ByRef str1 As String, ByRef str2 As String) As String
Attribute Concat3.VB_Description = "Concatenates three specified instances of String."
   Concat3 = this.BStringSingleton.Concat3(str0, str1, str2)
End Function

'@Static
'@Description("Concatenates four specified instances of String.")
Public Function Concat4(ByRef str0 As String, ByRef str1 As String, ByRef str2 As String, ByRef str3 As String) As String
Attribute Concat4.VB_Description = "Concatenates four specified instances of String."
   Concat4 = this.BStringSingleton.Concat4(str0, str1, str2, str3)
End Function

'@Static
'@Description("Concatenates the members of a constructed IEnumerable<T> collection of type String.")
Public Function Concat5(ByVal stringValues As mscorlib.IEnumerable) As String
Attribute Concat5.VB_Description = "Concatenates the members of a constructed IEnumerable<T> collection of type String."
   Concat5 = this.BStringSingleton.Concat5(stringValues)
End Function

'@Static
'@Description("Concatenates the string representations of the elements in a specified Object array.")
Public Function Concat6(ByRef values() As Variant) As String
Attribute Concat6.VB_Description = "Concatenates the string representations of the elements in a specified Object array."
   Concat6 = this.BStringSingleton.Concat6(values)
End Function

'@Static
'@Description("Creates the string representation of a specified object.")
Public Function Concat7(ByRef arg0 As Variant) As String
Attribute Concat7.VB_Description = "Creates the string representation of a specified object."
   Concat7 = this.BStringSingleton.Concat7(arg0)
End Function

'@Static
'@Description("Concatenates the string representations of two specified objects.")
Public Function Concat8(ByRef arg0 As Variant, ByRef arg1 As Variant) As String
Attribute Concat8.VB_Description = "Concatenates the string representations of two specified objects."
   Concat8 = this.BStringSingleton.Concat8(arg0, arg1)
End Function

'@Static
'@Description("Concatenates the string representations of three specified objects.")
Public Function Concat9(ByRef arg0 As Variant, ByRef arg1 As Variant, ByRef arg2 As Variant) As String
Attribute Concat9.VB_Description = "Concatenates the string representations of three specified objects."
   Concat9 = this.BStringSingleton.Concat9(arg0, arg1, arg2)
End Function

'@Static
'@Description("Concatenates the members of an IEnumerable<T> implementation.")
Public Function Concat10(ByVal values As mscorlib.IEnumerable) As String
Attribute Concat10.VB_Description = "Concatenates the members of an IEnumerable<T> implementation."
   Concat10 = this.BStringSingleton.Concat10(values)
End Function

'@Static
'@Description("Creates a new instance of String with the same value as a specified String.")
Public Function Copy(ByRef str As String) As String
Attribute Copy.VB_Description = "Creates a new instance of String with the same value as a specified String."
   Copy = this.BStringSingleton.Copy(str)
End Function

'@Static
'@Description("Determines whether two specified String objects have the same value.")
Public Function Equals(ByRef a As String, ByRef b As String) As Boolean
Attribute Equals.VB_Description = "Determines whether two specified String objects have the same value."
   Equals = this.BStringSingleton.Equals(a, b)
End Function

'@Static
'@Description("Determines whether two specified String objects have the same value. A parameter specifies the culture, case, and sort rules used in the comparison.")
Public Function Equals2(ByRef a As String, ByRef b As String, ByVal comparisonType As mscorlib.StringComparison) As Boolean
Attribute Equals2.VB_Description = "Determines whether two specified String objects have the same value. A parameter specifies the culture, case, and sort rules used in the comparison."
   Equals2 = this.BStringSingleton.Equals2(a, b, comparisonType)
End Function

'@Static
'@Description("Replaces the format item in a specified string with the string representation of a corresponding object in a specified array.")
Public Function Format(ByRef pFormat As String, ParamArray args() As Variant) As String
Attribute Format.VB_Description = "Replaces the format item in a specified string with the string representation of a corresponding object in a specified array."
    Dim params() As Variant
    params = args
    Format = this.BStringSingleton.Format(pFormat, params)
End Function

'@Static
'@Description("Replaces one or more format items in a string with the string representation of a specified object.")
Public Function Format2(ByRef pFormat As String, ByVal arg0 As Variant) As String
Attribute Format2.VB_Description = "Replaces one or more format items in a string with the string representation of a specified object."
   Format2 = this.BStringSingleton.Format2(pFormat, arg0)
End Function

'@Static
'@Description("Replaces the format items in a string with the string representation of two specified objects.")
Public Function Format3(ByRef pFormat As String, ByVal arg0 As Variant, ByVal arg1 As Variant) As String
Attribute Format3.VB_Description = "Replaces the format items in a string with the string representation of two specified objects."
   Format3 = this.BStringSingleton.Format3(pFormat, arg0, arg1)
End Function

'@Static
'@Description("Replaces the format items in a string with the string representation of three specified objects.")
Public Function Format4(ByRef pFormat As String, ByVal arg0 As Variant, ByVal arg1 As Variant, ByVal arg2 As Variant) As String
Attribute Format4.VB_Description = "Replaces the format items in a string with the string representation of three specified objects."
   Format4 = this.BStringSingleton.Format4(pFormat, arg0, arg1, arg2)
End Function

'@Static
'@Description("Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information.")
Public Function Format5(ByVal provider As mscorlib.IFormatProvider, ByRef pFormat As String, ParamArray args() As Variant) As String
Attribute Format5.VB_Description = "Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information."
    Dim params() As Variant
    params = args
    Format5 = this.BStringSingleton.Format5(provider, pFormat, params)
End Function

'@Description("Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information.")
Public Function Format6(ByVal provider As mscorlib.IFormatProvider, ByVal pFormat As String, ByVal arg0 As Variant) As String
Attribute Format6.VB_Description = "Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information."
   Format6 = this.BStringSingleton.Format6(provider, pFormat, arg0)
End Function

'@Description("Replaces the format items in a string with the string representation of two specified objects. A parameter supplies culture-specific formatting information.")
Public Function Format7(ByVal provider As mscorlib.IFormatProvider, ByVal pFormat As String, ByVal arg0 As Variant, ByVal arg1 As Variant) As String
Attribute Format7.VB_Description = "Replaces the format items in a string with the string representation of two specified objects. A parameter supplies culture-specific formatting information."
   Format7 = this.BStringSingleton.Format7(provider, pFormat, arg0, arg1)
End Function

'@Description("Replaces the format items in a string with the string representation of three specified objects. An parameter supplies culture-specific formatting information.")
Public Function Format8(ByVal provider As mscorlib.IFormatProvider, ByVal pFormat As String, ByVal arg0 As Variant, ByVal arg1 As Variant, ByVal arg2 As Variant) As String
Attribute Format8.VB_Description = "Replaces the format items in a string with the string representation of three specified objects. An parameter supplies culture-specific formatting information."
   Format8 = this.BStringSingleton.Format8(provider, pFormat, arg0, arg1, arg2)
End Function

'@Static
Friend Function Format9(ByRef pFormat As String, ByRef args() As Variant) As String
    Format9 = this.BStringSingleton.Format(pFormat, args)
End Function

'@Static
Friend Function Format10(ByVal provider As mscorlib.IFormatProvider, ByRef pFormat As String, ByRef args() As Variant) As String
    Format10 = this.BStringSingleton.Format5(provider, pFormat, args)
End Function

'@Static
'@Description("Indicates whether the specified string is null or an empty string ("").")
Public Function IsNullOrEmpty(ByRef value As String) As Boolean
Attribute IsNullOrEmpty.VB_Description = "Indicates whether the specified string is null or an empty string ("")."
   IsNullOrEmpty = this.BStringSingleton.IsNullOrEmpty(value)
End Function

'@Static
'@Description("Indicates whether a specified string is null, empty, or consists only of white-space characters.")
Public Function IsNullOrWhiteSpace(ByRef value As String) As Boolean
Attribute IsNullOrWhiteSpace.VB_Description = "Indicates whether a specified string is null, empty, or consists only of white-space characters."
   IsNullOrWhiteSpace = this.BStringSingleton.IsNullOrWhiteSpace(value)
End Function

'@Static
'@Description("Concatenates all the elements of a string array, using the specified separator between each element.")
Public Function Join(ByRef separator As String, ByRef value() As String) As String
Attribute Join.VB_Description = "Concatenates all the elements of a string array, using the specified separator between each element."
   Join = this.BStringSingleton.Join(separator, value)
End Function

'@Static
'@Description("Concatenates the elements of an object array, using the specified separator between each element.")
Public Function Join2(ByRef separator As String, ByRef values() As Variant) As String
Attribute Join2.VB_Description = "Concatenates the elements of an object array, using the specified separator between each element."
   Join2 = this.BStringSingleton.Join2(separator, values)
End Function

'@Static
'@Description("Concatenates the members of a constructed IEnumerable<T> collection of type String, using the specified separator between each member.")
Public Function Join3(ByRef separator As String, ByVal stringValues As mscorlib.IEnumerable) As String
Attribute Join3.VB_Description = "Concatenates the members of a constructed IEnumerable<T> collection of type String, using the specified separator between each member."
   Join3 = this.BStringSingleton.Join3(separator, stringValues)
End Function

'@Static
'@Description("Concatenates the specified elements of a string array, using the specified separator between each element.")
Public Function Join4(ByRef separator As String, ByRef value() As String, ByVal startIndex As Long, ByVal pCount As Long) As String
Attribute Join4.VB_Description = "Concatenates the specified elements of a string array, using the specified separator between each element."
   Join4 = this.BStringSingleton.Join4(separator, value, startIndex, pCount)
End Function

'--------------------------------------------------------------------
'Extension
'Added to get the type for arrays etc.
'--------------------------------------------------------------------

'@Static
'@Description("Returns a value indicating whether a specified substring occurs within this string, using the specified comparison rules.")
Public Function Contains(ByRef str As String, ByRef pSubstring As String, Optional ByVal comparisonType As mscorlib.StringComparison = StringComparison.StringComparison_Ordinal) As Boolean
Attribute Contains.VB_Description = "Returns a value indicating whether a specified substring occurs within this string, using the specified comparison rules."
   Contains = this.BStringSingleton.Contains(str, pSubstring, comparisonType)
End Function

'@Static
'@Description("Determines whether the end of this string instance matches the specified string when compared using the specified comparison option.")
Public Function EndsWith(ByRef str As String, ByRef pSubstring As String, Optional ByVal comparisonType As mscorlib.StringComparison = StringComparison.StringComparison_CurrentCulture) As Boolean
Attribute EndsWith.VB_Description = "Determines whether the end of this string instance matches the specified string when compared using the specified comparison option."
   EndsWith = this.BStringSingleton.EndsWith(str, pSubstring, comparisonType)
End Function

'@Static
'@Description("Determines whether the end of this string instance matches the specified string when compared using the specified culture.")
Public Function EndsWith2(ByRef str As String, ByRef pSubstring As String, ByVal ignoreCase As Boolean, ByVal culture As DotNetLib.CultureInfo) As Boolean
Attribute EndsWith2.VB_Description = "Determines whether the end of this string instance matches the specified string when compared using the specified culture."
   EndsWith2 = this.BStringSingleton.EndsWith2(str, pSubstring, ignoreCase, culture)
End Function

'@Static
'@Description("Splits a string into substrings based on specified delimiting characters and options.")
Public Function Split(ByRef str As String, ByRef separators As String, Optional ByVal options As DotNetLib.StringSplitOptions = StringSplitOptions_None) As String()
Attribute Split.VB_Description = "Splits a string into substrings based on specified delimiting characters and options."
   Split = this.BStringSingleton.Split(str, separators, options)
End Function

'@Static
'@Description("Splits a string into a maximum number of substrings based on specified delimiting characters and, optionally, options.")
Public Function Split2(ByRef str As String, ByVal separator As String, ByVal pCount As Long, Optional ByVal options As DotNetLib.StringSplitOptions = StringSplitOptions_None) As String()
Attribute Split2.VB_Description = "Splits a string into a maximum number of substrings based on specified delimiting characters and, optionally, options."
   Split2 = this.BStringSingleton.Split2(str, separator, pCount, options)
End Function

'@Static
'@Description("Splits a string into substrings based on a specified delimiting string and, optionally, options")
Public Function Split3(ByRef str As String, ByRef separator() As String, ByVal options As DotNetLib.StringSplitOptions) As String()
Attribute Split3.VB_Description = "Splits a string into substrings based on a specified delimiting string and, optionally, options"
   Split3 = this.BStringSingleton.Split3(str, separator, options)
End Function

'@Static
'@Description("Splits a string into a maximum number of substrings based on specified delimiting strings and, optionally, options.")
Public Function Split4(ByRef str As String, ByRef separator() As String, ByVal pCount As Long, ByVal options As DotNetLib.StringSplitOptions) As String()
Attribute Split4.VB_Description = "Splits a string into a maximum number of substrings based on specified delimiting strings and, optionally, options."
   Split4 = this.BStringSingleton.Split4(str, separator, pCount, options)
End Function

'@Static
'@Description("Determines whether the beginning of this string instance matches the specified string when compared using the specified comparison option.")
Public Function StartsWith(ByRef str As String, ByRef pSubstring As String, Optional ByVal comparisonType As mscorlib.StringComparison = StringComparison.StringComparison_CurrentCulture) As Boolean
Attribute StartsWith.VB_Description = "Determines whether the beginning of this string instance matches the specified string when compared using the specified comparison option."
   StartsWith = this.BStringSingleton.StartsWith(str, pSubstring, comparisonType)
End Function

'@Static
'@Description("Determines whether the beginning of this string instance matches the specified string when compared using the specified culture.")
Public Function StartsWith2(ByRef str As String, ByRef pSubstring As String, ByVal ignoreCase As Boolean, ByVal culture As DotNetLib.CultureInfo) As Boolean
Attribute StartsWith2.VB_Description = "Determines whether the beginning of this string instance matches the specified string when compared using the specified culture."
   StartsWith2 = this.BStringSingleton.StartsWith2(str, pSubstring, ignoreCase, culture)
End Function

'@Static
'@DefaultMember
'@Description("Gets the String Type.")
Public Function GetType() As DotNetLib.Type
Attribute GetType.VB_Description = "Gets the String Type."
Attribute GetType.VB_UserMemId = 0
    Set GetType = this.StringType
End Function

'@Static
'@Description("Initializes a new instance a string to the value indicated by an string of Unicode characters, converting any escaped characters in the input string.")
Public Function Unescape(ByRef value As String) As String
Attribute Unescape.VB_Description = "Initializes a new instance a string to the value indicated by an string of Unicode characters, converting any escaped characters in the input string."
   Unescape = this.BStringSingleton.Unescape(value)
End Function


'@Static
'@Description("")
Public Function IndexOf(ByRef pValue As String, ByRef pSubstring As String) As Long
   IndexOf = InStr(pValue, pSubstring) - 1
End Function

'@Static
'@Description("")
Public Function IndexOf2(ByRef pValue As String, ByRef pSubstring As String, ByVal compareMode As Long) As Long
   IndexOf2 = InStr(pValue, pSubstring, compareMode) - 1
End Function

''@Static
''@Description("")
'Public Function IndexOf3(ByRef pValue As String, ByRef pSubstring As String, pComparison As mscorlib.StringComparison) As Long
'End Function
