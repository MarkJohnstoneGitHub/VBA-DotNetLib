VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeSpan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "TimeSpan factory methods and static members."
'Rubberduck annotations
'@Folder "VBADotNetLib.System"
'@PredeclaredId
'@ModuleDescription "TimeSpan factory methods and static members."

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v2.0 August 4, 2023
'@LastModified August 15, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8.1

'@Remarks
' Singleton class for TimeSpan containing constructors and static members.

Option Explicit

Private Type TTimeSpan
    TimeSpanSingleton As DotNetLib.TimeSpanSingleton
End Type

Private this As TTimeSpan

Private Sub Class_Initialize()
    Set this.TimeSpanSingleton = New DotNetLib.TimeSpanSingleton
End Sub

Private Sub Class_Terminate()
    Set this.TimeSpanSingleton = Nothing
End Sub

'--------------------------------------------------------------------
'Constructors
'--------------------------------------------------------------------

'@Static
'@Description("Initializes a new instance of the TimeSpan structure to the specified number of ticks.")
Public Function CreateFromTicks(ByVal pTicks As LongLong) As DotNetLib.ITimeSpan
Attribute CreateFromTicks.VB_Description = "Initializes a new instance of the TimeSpan structure to the specified number of ticks."
    Set CreateFromTicks = this.TimeSpanSingleton.CreateFromTicks(pTicks)
End Function

'@Static
'@Description("Initializes a new instance of the TimeSpan structure to a specified number of hours, minutes, and seconds.")
Public Function Create(ByVal pHours As Long, ByVal pMinutes As Long, ByVal pSeconds As Long) As DotNetLib.ITimeSpan
Attribute Create.VB_Description = "Initializes a new instance of the TimeSpan structure to a specified number of hours, minutes, and seconds."
    Set Create = this.TimeSpanSingleton.Create(pHours, pMinutes, pSeconds)
End Function

'@Static
'@Description("Initializes a new instance of the TimeSpan structure to a specified number of days, hours, minutes, and seconds.")
Public Function Create2(ByVal pDays As Long, ByVal pHours As Long, ByVal pMinutes As Long, ByVal pSeconds As Long) As DotNetLib.ITimeSpan
Attribute Create2.VB_Description = "Initializes a new instance of the TimeSpan structure to a specified number of days, hours, minutes, and seconds."
    Set Create2 = this.TimeSpanSingleton.Create2(pDays, pHours, pMinutes, pSeconds)
End Function

'@Static
'@Description("Initializes a new instance of the TimeSpan structure to a specified number of days, hours, minutes, seconds, and milliseconds.")
Public Function Create3(ByVal pDays As Long, ByVal pHours As Long, ByVal pMinutes As Long, ByVal pSeconds As Long, ByVal pMilliseconds As Long) As DotNetLib.ITimeSpan
Attribute Create3.VB_Description = "Initializes a new instance of the TimeSpan structure to a specified number of days, hours, minutes, seconds, and milliseconds."
    Set Create3 = this.TimeSpanSingleton.Create3(pDays, pHours, pMinutes, pSeconds, pMilliseconds)
End Function

'--------------------------------------------------------------------
'Fields
'--------------------------------------------------------------------

'@Static
'@Description("Represents the maximum TimeSpan value. This field is read-only.")
Public Property Get MaxValue() As DotNetLib.ITimeSpan
Attribute MaxValue.VB_Description = "Represents the maximum TimeSpan value. This field is read-only."
    Set MaxValue = this.TimeSpanSingleton.MaxValue
End Property

'@Static
'@Description("Represents the minimum TimeSpan value. This field is read-only.")
Public Property Get MinValue() As DotNetLib.ITimeSpan
Attribute MinValue.VB_Description = "Represents the minimum TimeSpan value. This field is read-only."
    Set MinValue = this.TimeSpanSingleton.MinValue
End Property

'@Static
'@Description("Represents the number of ticks in 1 day. This field is constant.")
Public Property Get TicksPerDay() As LongLong
Attribute TicksPerDay.VB_Description = "Represents the number of ticks in 1 day. This field is constant."
    TicksPerDay = this.TimeSpanSingleton.TicksPerDay
End Property

'@Static
'@Description("Represents the number of ticks in 1 hour. This field is constant.")
Public Property Get TicksPerHour() As LongLong
Attribute TicksPerHour.VB_Description = "Represents the number of ticks in 1 hour. This field is constant."
    TicksPerHour = this.TimeSpanSingleton.TicksPerHour()
End Property

'@Static
'@Description("Represents the number of ticks in 1 millisecond. This field is constant.")
Public Property Get TicksPerMillisecond() As LongLong
Attribute TicksPerMillisecond.VB_Description = "Represents the number of ticks in 1 millisecond. This field is constant."
    TicksPerMillisecond = this.TimeSpanSingleton.TicksPerMillisecond()
End Property

'@Static
'@Description("Represents the number of ticks in 1 minute. This field is constant.")
Public Property Get TicksPerMinute() As LongLong
Attribute TicksPerMinute.VB_Description = "Represents the number of ticks in 1 minute. This field is constant."
    TicksPerMinute = this.TimeSpanSingleton.TicksPerMinute()
End Property

'@Static
'@Description("Represents the number of ticks in 1 second.")
Public Property Get TicksPerSecond() As LongLong
Attribute TicksPerSecond.VB_Description = "Represents the number of ticks in 1 second."
    TicksPerSecond = this.TimeSpanSingleton.TicksPerSecond
End Property

'@Static
'@Description("Represents the zero TimeSpan value. This field is read-only.")
Public Property Get Zero() As DotNetLib.ITimeSpan
Attribute Zero.VB_Description = "Represents the zero TimeSpan value. This field is read-only."
    Set Zero = this.TimeSpanSingleton.Zero
End Property

'--------------------------------------------------------------------
'Methods
'--------------------------------------------------------------------

'@Static
'@Description("Compares two TimeSpan values and returns an integer that indicates whether the first value is shorter than, equal to, or longer than the second value.")
Public Function Compare(ByVal t1 As DotNetLib.TimeSpan, ByVal t2 As DotNetLib.TimeSpan) As Long
Attribute Compare.VB_Description = "Compares two TimeSpan values and returns an integer that indicates whether the first value is shorter than, equal to, or longer than the second value."
    Compare = this.TimeSpanSingleton.Compare(t1, t2)
End Function

'@Static
'@Description("Returns a value that indicates whether two specified instances of TimeSpan are equal.")
Public Function Equals(ByVal t1 As DotNetLib.TimeSpan, ByVal t2 As DotNetLib.TimeSpan) As Boolean
Attribute Equals.VB_Description = "Returns a value that indicates whether two specified instances of TimeSpan are equal."
    Equals = this.TimeSpanSingleton.Equals(t1, t2)
End Function

'@Static
'@Description("Returns a TimeSpan that represents a specified number of days, where the specification is accurate to the nearest millisecond.")
Public Function FromDays(ByVal value As Double) As DotNetLib.ITimeSpan
Attribute FromDays.VB_Description = "Returns a TimeSpan that represents a specified number of days, where the specification is accurate to the nearest millisecond."
    Set FromDays = this.TimeSpanSingleton.FromDays(value)
End Function

'@Static
'@Description("Returns a TimeSpan that represents a specified number of hours, where the specification is accurate to the nearest millisecond.")
Public Function FromHours(ByVal value As Double) As DotNetLib.ITimeSpan
Attribute FromHours.VB_Description = "Returns a TimeSpan that represents a specified number of hours, where the specification is accurate to the nearest millisecond."
    Set FromHours = this.TimeSpanSingleton.FromHours(value)
End Function

'@Static
'@Description("Returns a TimeSpan that represents a specified number of milliseconds.")
Public Function FromMilliseconds(ByVal value As Double) As DotNetLib.ITimeSpan
Attribute FromMilliseconds.VB_Description = "Returns a TimeSpan that represents a specified number of milliseconds."
    Set FromMilliseconds = this.TimeSpanSingleton.FromMilliseconds(value)
End Function

'@Static
'@Description("Returns a TimeSpan that represents a specified number of minutes, where the specification is accurate to the nearest millisecond.")
Public Function FromMinutes(ByVal value As Double) As DotNetLib.ITimeSpan
Attribute FromMinutes.VB_Description = "Returns a TimeSpan that represents a specified number of minutes, where the specification is accurate to the nearest millisecond."
    Set FromMinutes = this.TimeSpanSingleton.FromMinutes(value)
End Function

'@Static
'@Description("Returns a TimeSpan that represents a specified number of seconds, where the specification is accurate to the nearest millisecond.")
Public Function FromSeconds(ByVal value As Double) As DotNetLib.ITimeSpan
Attribute FromSeconds.VB_Description = "Returns a TimeSpan that represents a specified number of seconds, where the specification is accurate to the nearest millisecond."
    Set FromSeconds = this.TimeSpanSingleton.FromSeconds(value)
End Function

'@Static
'@Description("Returns a TimeSpan that represents a specified time, where the specification is in units of ticks.")
Public Function FromTicks(ByVal value As LongLong) As DotNetLib.ITimeSpan
Attribute FromTicks.VB_Description = "Returns a TimeSpan that represents a specified time, where the specification is in units of ticks."
    Set FromTicks = this.TimeSpanSingleton.FromTicks(value)
End Function

'@Static
'@Description("Converts the string representation of a time interval to its TimeSpan equivalent.")
Public Function Parse(ByVal s As String) As DotNetLib.ITimeSpan
Attribute Parse.VB_Description = "Converts the string representation of a time interval to its TimeSpan equivalent."
    Set Parse = this.TimeSpanSingleton.Parse(s)
End Function

'@Static
'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified culture-specific format information.")
Public Function Parse2(ByVal strInput As String, ByVal formatProvider As IFormatProvider) As DotNetLib.ITimeSpan
Attribute Parse2.VB_Description = "Converts the string representation of a time interval to its TimeSpan equivalent by using the specified culture-specific format information."
    Set Parse2 = this.TimeSpanSingleton.Parse2(strInput, formatProvider)
End Function

'@Static
'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")
Public Function ParseExact(ByVal strInput As String, ByVal Format As String, ByVal formatProvider As IFormatProvider) As DotNetLib.ITimeSpan
Attribute ParseExact.VB_Description = "Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly."
    Set ParseExact = this.TimeSpanSingleton.ParseExact(strInput, Format, formatProvider)
End Function

'@Static
'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified array of format strings and culture-specific format information. The format of the string representation must match one of the specified formats exactly.")
Public Function ParseExact2(ByVal strInput As String, ByRef formats() As String, ByVal formatProvider As IFormatProvider) As DotNetLib.ITimeSpan
Attribute ParseExact2.VB_Description = "Converts the string representation of a time interval to its TimeSpan equivalent by using the specified array of format strings and culture-specific format information. The format of the string representation must match one of the specified formats exactly."
    Set ParseExact2 = this.TimeSpanSingleton.ParseExact2(strInput, formats, formatProvider)
End Function

'@Static
'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format, culture-specific format information, and styles. The format of the string representation must match the specified format exactly.")
Public Function ParseExact3(ByVal strInput As String, ByVal Format As String, ByVal formatProvider As IFormatProvider, ByVal styles As DotNetLib.TimeSpanStyles) As DotNetLib.ITimeSpan
Attribute ParseExact3.VB_Description = "Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format, culture-specific format information, and styles. The format of the string representation must match the specified format exactly."
    Set ParseExact3 = this.TimeSpanSingleton.ParseExact3(strInput, Format, formatProvider, styles)
End Function

'@Static
'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified formats, culture-specific format information, and styles. The format of the string representation must match one of the specified formats exactly")
Public Function ParseExact4(ByVal strInput As String, ByRef formats() As String, ByVal formatProvider As IFormatProvider, ByVal styles As DotNetLib.TimeSpanStyles) As DotNetLib.ITimeSpan
Attribute ParseExact4.VB_Description = "Converts the string representation of a time interval to its TimeSpan equivalent by using the specified formats, culture-specific format information, and styles. The format of the string representation must match one of the specified formats exactly"
    Set ParseExact4 = this.TimeSpanSingleton.ParseExact4(strInput, formats, formatProvider, styles)
End Function

'@Static
'@Description("Converts the string representation of a time interval to its TimeSpan equivalent and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse(ByVal s As String, ByRef Result As DotNetLib.TimeSpan) As Boolean
Attribute TryParse.VB_Description = "Converts the string representation of a time interval to its TimeSpan equivalent and returns a value that indicates whether the conversion succeeded."
    TryParse = this.TimeSpanSingleton.TryParse(s, Result)
End Function

'@Static
'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified culture-specific formatting information, and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse2(ByVal strInput As String, ByVal formatProvider As IFormatProvider, ByRef Result As DotNetLib.TimeSpan) As Boolean
Attribute TryParse2.VB_Description = "Converts the string representation of a time interval to its TimeSpan equivalent by using the specified culture-specific formatting information, and returns a value that indicates whether the conversion succeeded."
    TryParse2 = this.TimeSpanSingleton.TryParse2(strInput, formatProvider, Result)
End Function

'@Static
'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")
Public Function TryParseExact(ByVal strInput As String, ByVal Format As String, ByVal formatProvider As IFormatProvider, ByRef Result As DotNetLib.TimeSpan) As Boolean
Attribute TryParseExact.VB_Description = "Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly."
    TryParseExact = this.TimeSpanSingleton.TryParseExact(strInput, Format, formatProvider, Result)
End Function

'@Static
'@Description("Converts the specified string representation of a time interval to its TimeSpan equivalent by using the specified formats and culture-specific format information. The format of the string representation must match one of the specified formats exactly.")
Public Function TryParseExact2(ByVal strInput As String, ByRef formats() As String, ByVal formatProvider As IFormatProvider, ByRef Result As DotNetLib.TimeSpan) As Boolean
Attribute TryParseExact2.VB_Description = "Converts the specified string representation of a time interval to its TimeSpan equivalent by using the specified formats and culture-specific format information. The format of the string representation must match one of the specified formats exactly."
    TryParseExact2 = this.TimeSpanSingleton.TryParseExact2(strInput, formats, formatProvider, Result)
End Function

'@Static
'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format, culture-specific format information and styles. The format of the string representation must match the specified format exactly.")
Public Function TryParseExact3(ByVal strInput As String, ByVal Format As String, ByVal formatProvider As IFormatProvider, ByVal styles As DotNetLib.TimeSpanStyles, ByRef Result As DotNetLib.TimeSpan) As Boolean
Attribute TryParseExact3.VB_Description = "Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format, culture-specific format information and styles. The format of the string representation must match the specified format exactly."
    TryParseExact3 = this.TimeSpanSingleton.TryParseExact3(strInput, Format, formatProvider, styles, Result)
End Function

'@Static
'@Description("Converts the specified string representation of a time interval to its TimeSpan equivalent by using the specified formats, culture-specific format information and styles. The format of the string representation must match one of the specified formats exactly.")
Public Function TryParseExact4(ByVal strInput As String, ByRef formats() As String, ByVal formatProvider As IFormatProvider, ByVal styles As DotNetLib.TimeSpanStyles, ByRef Result As DotNetLib.TimeSpan) As Boolean
Attribute TryParseExact4.VB_Description = "Converts the specified string representation of a time interval to its TimeSpan equivalent by using the specified formats, culture-specific format information and styles. The format of the string representation must match one of the specified formats exactly."
    TryParseExact4 = this.TimeSpanSingleton.TryParseExact4(strInput, formats, formatProvider, styles, Result)
End Function

'--------------------------------------------------------------------
'Operators
'--------------------------------------------------------------------

'@Static
'@Description("Adds two specified TimeSpan instances.")
Public Function Addition(ByVal t1 As DotNetLib.TimeSpan, ByVal t2 As DotNetLib.TimeSpan) As DotNetLib.ITimeSpan
Attribute Addition.VB_Description = "Adds two specified TimeSpan instances."
    Set Addition = this.TimeSpanSingleton.Addition(t1, t2)
End Function

'@Static
'@Description("Indicates whether two TimeSpan instances are equal.")
Public Function Equality(ByVal t1 As DotNetLib.TimeSpan, ByVal t2 As DotNetLib.TimeSpan) As Boolean
Attribute Equality.VB_Description = "Indicates whether two TimeSpan instances are equal."
    Equality = this.TimeSpanSingleton.Equality(t1, t2)
End Function

'@Static
'@Description("Indicates whether a specified TimeSpan is greater than another specified TimeSpan.")
Public Function GreaterThan(ByVal t1 As DotNetLib.TimeSpan, ByVal t2 As DotNetLib.TimeSpan) As Boolean
Attribute GreaterThan.VB_Description = "Indicates whether a specified TimeSpan is greater than another specified TimeSpan."
    GreaterThan = this.TimeSpanSingleton.GreaterThan(t1, t2)
End Function

'@Static
'@Description("Indicates whether a specified TimeSpan is greater than or equal to another specified TimeSpan.")
Public Function GreaterThanOrEqual(ByVal t1 As DotNetLib.TimeSpan, ByVal t2 As DotNetLib.TimeSpan) As Boolean
Attribute GreaterThanOrEqual.VB_Description = "Indicates whether a specified TimeSpan is greater than or equal to another specified TimeSpan."
    GreaterThanOrEqual = this.TimeSpanSingleton.GreaterThanOrEqual(t1, t2)
End Function

'@Static
'@Description("Indicates whether two TimeSpan instances are not equal.")
Public Function Inequality(ByVal t1 As DotNetLib.TimeSpan, ByVal t2 As DotNetLib.TimeSpan) As Boolean
Attribute Inequality.VB_Description = "Indicates whether two TimeSpan instances are not equal."
    Inequality = this.TimeSpanSingleton.Inequality(t1, t2)
End Function

'@Static
'@Description("Indicates whether a specified TimeSpan is less than another specified TimeSpan.")
Public Function LessThan(ByVal t1 As DotNetLib.TimeSpan, ByVal t2 As DotNetLib.TimeSpan) As Boolean
Attribute LessThan.VB_Description = "Indicates whether a specified TimeSpan is less than another specified TimeSpan."
    LessThan = this.TimeSpanSingleton.LessThan(t1, t2)
End Function

'@Static
'@Description("Indicates whether a specified TimeSpan is less than or equal to another specified TimeSpan.")
Public Function LessThanOrEqual(ByVal t1 As DotNetLib.TimeSpan, ByVal t2 As DotNetLib.TimeSpan) As Boolean
Attribute LessThanOrEqual.VB_Description = "Indicates whether a specified TimeSpan is less than or equal to another specified TimeSpan."
    LessThanOrEqual = this.TimeSpanSingleton.LessThanOrEqual(t1, t2)
End Function

'@Static
'@Description("Subtracts a specified TimeSpan from another specified TimeSpan.")
Public Function Subtraction(ByVal t1 As DotNetLib.TimeSpan, ByVal t2 As DotNetLib.TimeSpan) As DotNetLib.ITimeSpan
Attribute Subtraction.VB_Description = "Subtracts a specified TimeSpan from another specified TimeSpan."
    Set Subtraction = this.TimeSpanSingleton.Subtraction(t1, t2)
End Function

'@Static
'@Description("Returns a TimeSpan whose value is the negated value of the specified instance.")
Public Function UnaryNegation(ByVal ts As DotNetLib.TimeSpan) As DotNetLib.ITimeSpan
Attribute UnaryNegation.VB_Description = "Returns a TimeSpan whose value is the negated value of the specified instance."
    Set UnaryNegation = this.TimeSpanSingleton.UnaryNegation(ts)
End Function

'@Static
'@Description("Returns the specified instance of TimeSpan.")
Public Function UnaryPlus(ByVal ts As DotNetLib.TimeSpan) As DotNetLib.ITimeSpan
Attribute UnaryPlus.VB_Description = "Returns the specified instance of TimeSpan."
    Set UnaryPlus = this.TimeSpanSingleton.UnaryPlus(ts)
End Function
