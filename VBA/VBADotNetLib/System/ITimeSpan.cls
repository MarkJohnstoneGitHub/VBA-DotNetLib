VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITimeSpan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "VBADotNetLib.System"
'@Interface

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 29, 2023
'@LastModified July 29, 2023

Option Explicit

'@Description("Gets the days component of the time interval represented by the current TimeSpan structure.")
Public Property Get Days() As Long
Attribute Days.VB_Description = "Gets the days component of the time interval represented by the current TimeSpan structure."
End Property

'@Description("Gets the hours component of the time interval represented by the current TimeSpan structure.")
Public Property Get Hours() As Long
Attribute Hours.VB_Description = "Gets the hours component of the time interval represented by the current TimeSpan structure."
End Property

'@Description("Gets the milliseconds component of the time interval represented by the current TimeSpan structure.")
Public Property Get Milliseconds() As Long
Attribute Milliseconds.VB_Description = "Gets the milliseconds component of the time interval represented by the current TimeSpan structure."
End Property

'@Description("Gets the minutes component of the time interval represented by the current TimeSpan structure.")
Public Property Get Minutes() As Long
Attribute Minutes.VB_Description = "Gets the minutes component of the time interval represented by the current TimeSpan structure."
End Property

'@Description("Gets the seconds component of the time interval represented by the current TimeSpan structure.")
Public Property Get Seconds() As Long
Attribute Seconds.VB_Description = "Gets the seconds component of the time interval represented by the current TimeSpan structure."
End Property

'@Description("Gets the number of ticks that represent the value of the current TimeSpan structure.")
Public Property Get Ticks() As LongLong
Attribute Ticks.VB_Description = "Gets the number of ticks that represent the value of the current TimeSpan structure."
End Property

'@Description("Gets the value of the current TimeSpan structure expressed in whole and fractional days.")
Public Property Get TotalDays() As Double
Attribute TotalDays.VB_Description = "Gets the value of the current TimeSpan structure expressed in whole and fractional days."
End Property

'@Description("Gets the value of the current TimeSpan structure expressed in whole and fractional hours.")
Public Property Get TotalHours() As Double
Attribute TotalHours.VB_Description = "Gets the value of the current TimeSpan structure expressed in whole and fractional hours."
End Property

'@Description("Gets the value of the current TimeSpan structure expressed in whole and fractional milliseconds.")
Public Property Get TotalMilliseconds() As Double
Attribute TotalMilliseconds.VB_Description = "Gets the value of the current TimeSpan structure expressed in whole and fractional milliseconds."
End Property

'@Description("Gets the value of the current TimeSpan structure expressed in whole and fractional minutes.")
Public Property Get TotalMinutes() As Double
Attribute TotalMinutes.VB_Description = "Gets the value of the current TimeSpan structure expressed in whole and fractional minutes."
End Property

'@Description("Gets the value of the current TimeSpan structure expressed in whole and fractional seconds.")
Public Property Get TotalSeconds() As Double
Attribute TotalSeconds.VB_Description = "Gets the value of the current TimeSpan structure expressed in whole and fractional seconds."
End Property

'@Description("Returns a new TimeSpan object whose value is the sum of the specified TimeSpan object and this instance.")
Public Function Add(ByVal ts As TimeSpan) As TimeSpan
Attribute Add.VB_Description = "Returns a new TimeSpan object whose value is the sum of the specified TimeSpan object and this instance."
End Function

'@Description("Compares this instance to a specified TimeSpan object and returns an integer that indicates whether this instance is shorter than, equal to, or longer than the TimeSpan object.")
Public Function CompareTo(ByVal value As TimeSpan) As Long
Attribute CompareTo.VB_Description = "Compares this instance to a specified TimeSpan object and returns an integer that indicates whether this instance is shorter than, equal to, or longer than the TimeSpan object."
End Function

'@Description("Returns a new TimeSpan object whose value is the absolute value of the current TimeSpan object.")
Public Function Duration() As TimeSpan
Attribute Duration.VB_Description = "Returns a new TimeSpan object whose value is the absolute value of the current TimeSpan object."
End Function

'@Description("Returns a value indicating whether this instance is equal to a specified TimeSpan object.")
Public Function Equals(ByVal value As TimeSpan) As Boolean
Attribute Equals.VB_Description = "Returns a value indicating whether this instance is equal to a specified TimeSpan object."
End Function

'@Description("Returns a hash code for this instance.")
Public Function GetHashCode() As Long
Attribute GetHashCode.VB_Description = "Returns a hash code for this instance."
End Function

'@Description("Returns a new TimeSpan object whose value is the negated value of this instance.")
Public Function Negate() As TimeSpan
Attribute Negate.VB_Description = "Returns a new TimeSpan object whose value is the negated value of this instance."
End Function

'@Description("Returns a new TimeSpan object whose value is the difference between the specified TimeSpan object and this instance.")
Public Function Subtract(ByVal ts As TimeSpan) As TimeSpan
Attribute Subtract.VB_Description = "Returns a new TimeSpan object whose value is the difference between the specified TimeSpan object and this instance."
End Function

'@Description("Converts the value of the current TimeSpan object to its equivalent string representation.")
Public Property Get ToString() As String
Attribute ToString.VB_Description = "Converts the value of the current TimeSpan object to its equivalent string representation."
End Property

'@Description("Converts the value of the current TimeSpan object to its equivalent string representation by using the specified format.")
Public Function ToString2(ByVal Format As String) As String
Attribute ToString2.VB_Description = "Converts the value of the current TimeSpan object to its equivalent string representation by using the specified format."
End Function

''@Description("Converts the value of the current TimeSpan object to its equivalent string representation by using the specified format and culture-specific formatting information.")
Public Function ToString3(ByVal Format As String, ByVal formatProvider As IFormatProvider) As String
End Function
