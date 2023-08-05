VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents an instant in time, typically expressed as a date and time of day."
'Rubberduck annotations
'@Folder "VBADotNetLib.System"
'@PredeclaredId
'@ModuleDescription("Represents an instant in time, typically expressed as a date and time of day.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v2.0 August 4, 2023
'@LastModified August 4, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1

'@Remarks
' Singleton class for DateTime containing constructors and static members.

Option Explicit

Private Type TDateTime
    DateTimeSingleton As DotNetLib.IDateTimeSingleton
End Type

Private this As TDateTime

Private Sub Class_Initialize()
    Set this.DateTimeSingleton = New DotNetLib.DateTime
End Sub

Private Sub Class_Terminate()
    Set this.DateTimeSingleton = Nothing
End Sub

'--------------------------------------------------------------------
'Constructors
'--------------------------------------------------------------------

'@Static
'@Description("Initializes a new instance of the DateTime structure to a specified number of ticks and to Coordinated Universal Time (UTC) or local time.")
Public Function CreateFromTicks(ByVal pTicks As LongLong, Optional ByVal pKind As DotNetLib.DateTimeKind = DateTimeKind_Unspecified) As DotNetLib.IDateTime
Attribute CreateFromTicks.VB_Description = "Initializes a new instance of the DateTime structure to a specified number of ticks and to Coordinated Universal Time (UTC) or local time."
    Set CreateFromTicks = this.DateTimeSingleton.CreateFromTicks(pTicks, pKind)
End Function

'@Static
'@Description("Initializes a new instance of the DateTime structure to the specified year, month, and day.")
Public Function CreateFromDate(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long) As DotNetLib.IDateTime
Attribute CreateFromDate.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, and day."
    Set CreateFromDate = this.DateTimeSingleton.CreateFromDate(pYear, pMonth, pDay)
End Function

'@Static
'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond.")
Public Function CreateFromDateTime(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long, Optional ByVal pMillisecond As Long = 0) As DotNetLib.IDateTime
Attribute CreateFromDateTime.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond."
    Set CreateFromDateTime = this.DateTimeSingleton.CreateFromDateTime(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond)
End Function

'@Static
'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and Coordinated Universal Time (UTC) or local time.")
Public Function CreateFromDateTimeKind(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long, ByVal pKind As DotNetLib.DateTimeKind) As DotNetLib.IDateTime
Attribute CreateFromDateTimeKind.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and Coordinated Universal Time (UTC) or local time."
    Set CreateFromDateTimeKind = this.DateTimeSingleton.CreateFromDateTimeKind(pYear, pMonth, pDay, pHour, pMinute, pSecond, pKind)
End Function

'@Static
'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time.")
Public Function CreateFromDateTimeKind2(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long, ByVal pMillisecond As Long, ByVal pKind As DotNetLib.DateTimeKind) As DotNetLib.IDateTime
Attribute CreateFromDateTimeKind2.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time."
    Set CreateFromDateTimeKind2 = this.DateTimeSingleton.CreateFromDateTimeKind2(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond, pKind)
End Function

'--------------------------------------------------------------------
'Fields
'--------------------------------------------------------------------

'@Static
'@Description("Represents the largest possible value of DateTime. This field is read-only.")
Public Property Get MaxValue() As DotNetLib.IDateTime
Attribute MaxValue.VB_Description = "Represents the largest possible value of DateTime. This field is read-only."
    Set MaxValue = this.DateTimeSingleton.MaxValue
End Property

'@Static
'@Description("Represents the smallest possible value of DateTime. This field is read-only.")
Public Property Get MinValue() As DotNetLib.IDateTime
Attribute MinValue.VB_Description = "Represents the smallest possible value of DateTime. This field is read-only."
    Set MinValue = this.DateTimeSingleton.MinValue
End Property

'--------------------------------------------------------------------
'Properties
'--------------------------------------------------------------------
'@Static
'@Description("Gets a DateTime object that is set to the current date and time on this computer, expressed as the local time.")
Public Property Get Now() As DotNetLib.IDateTime
Attribute Now.VB_Description = "Gets a DateTime object that is set to the current date and time on this computer, expressed as the local time."
    Set Now = this.DateTimeSingleton.Now()
End Property

'@Static
'@Description("Gets the current date.")
Public Property Get Today() As DotNetLib.IDateTime
Attribute Today.VB_Description = "Gets the current date."
    Set Today = this.DateTimeSingleton.Today()
End Property

'@Static
'@Description("Gets a DateTime object that is set to the current date and time on this computer, expressed as the Coordinated Universal Time (UTC).")
Public Property Get UtcNow() As DotNetLib.IDateTime
Attribute UtcNow.VB_Description = "Gets a DateTime object that is set to the current date and time on this computer, expressed as the Coordinated Universal Time (UTC)."
    Set UtcNow = this.DateTimeSingleton.UtcNow()
End Property

'--------------------------------------------------------------------
'Methods
'--------------------------------------------------------------------

'@Static
'@Description("Compares two instances of DateTime and returns an integer that indicates whether the first instance is earlier than, the same as, or later than the second instance.")
Public Function Compare(ByVal t1 As DotNetLib.DateTime, ByVal t2 As DotNetLib.DateTime) As Long
Attribute Compare.VB_Description = "Compares two instances of DateTime and returns an integer that indicates whether the first instance is earlier than, the same as, or later than the second instance."
    Compare = this.DateTimeSingleton.Compare(t1, t2)
End Function

'@Static
'@Description("Returns the number of days in the specified month and year.")
Public Function DaysInMonth(ByVal pYear As Long, ByVal pMonth As Long) As Long
Attribute DaysInMonth.VB_Description = "Returns the number of days in the specified month and year."
    DaysInMonth = this.DateTimeSingleton.DaysInMonth(pYear, pMonth)
End Function

'@Static
'@Description("Returns a value indicating whether two DateTime instances have the same date and time value.")
Public Function Equals3(ByVal t1 As DotNetLib.DateTime, ByVal t2 As DotNetLib.DateTime) As Boolean
Attribute Equals3.VB_Description = "Returns a value indicating whether two DateTime instances have the same date and time value."
    Equals3 = this.DateTimeSingleton.Equals3(t1, t2)
End Function

'@Static
'@Description("Deserializes a 64-bit binary value and recreates an original serialized DateTime object.")
Public Function FromBinary(ByVal dateData As LongLong) As DotNetLib.IDateTime
Attribute FromBinary.VB_Description = "Deserializes a 64-bit binary value and recreates an original serialized DateTime object."
    Set FromBinary = this.DateTimeSingleton.FromBinary(dateData)
End Function

'@Static
'@Description("Converts the specified Windows file time to an equivalent local time.")
Public Function FromFileTime(ByVal fileTime As LongLong) As DotNetLib.IDateTime
Attribute FromFileTime.VB_Description = "Converts the specified Windows file time to an equivalent local time."
    Set FromFileTime = this.DateTimeSingleton.FromFileTime(fileTime)
End Function

'@Static
'@Description("Converts the specified Windows file time to an equivalent UTC time.")
Public Function FromFileTimeUtc(ByVal fileTime As LongLong) As DotNetLib.IDateTime
Attribute FromFileTimeUtc.VB_Description = "Converts the specified Windows file time to an equivalent UTC time."
    Set FromFileTimeUtc = this.DateTimeSingleton.FromFileTimeUtc(fileTime)
End Function

'@Static
'@Description("Returns a DateTime equivalent to the specified OLE Automation Date.")
Public Function FromOADate(ByVal oaDate As Date) As DotNetLib.IDateTime
Attribute FromOADate.VB_Description = "Returns a DateTime equivalent to the specified OLE Automation Date."
    Set FromOADate = this.DateTimeSingleton.FromOADate(oaDate)
End Function

'@Static
'@Description("Returns an indication whether the specified year is a leap year.")
Public Function IsLeapYear(ByVal pYear As Long) As Boolean
Attribute IsLeapYear.VB_Description = "Returns an indication whether the specified year is a leap year."
    IsLeapYear = this.DateTimeSingleton.IsLeapYear(pYear)
End Function

'@Static
'@Description("Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture.")
Public Function Parse(ByVal s As String) As DotNetLib.IDateTime
Attribute Parse.VB_Description = "Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture."
    Set Parse = this.DateTimeSingleton.Parse(s)
End Function

'@Static
'@Description("Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture.")
Public Function Parse2(ByVal s As String, ByVal provider As IFormatProvider) As DotNetLib.IDateTime
Attribute Parse2.VB_Description = "Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture."
    Set Parse2 = this.DateTimeSingleton.Parse2(s, provider)
End Function

'@Static
'@Description("Defines the formatting options that customize string parsing for some date and time parsing methods.  This enumeration supports a bitwise combination of its member values.")
Public Function Parse3(ByVal s As String, ByVal provider As IFormatProvider, ByVal styles As DateTimeStyles) As DotNetLib.IDateTime
Attribute Parse3.VB_Description = "Defines the formatting options that customize string parsing for some date and time parsing methods.  This enumeration supports a bitwise combination of its member values."
    Set Parse3 = this.DateTimeSingleton.Parse3(s, provider, styles)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")
Public Function ParseExact(ByVal s As String, ByVal pFormat As String, ByVal provider As IFormatProvider) As DotNetLib.IDateTime
Attribute ParseExact.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly."
    Set ParseExact = this.DateTimeSingleton.ParseExact(s, pFormat, provider)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly or an exception is thrown.")
Public Function ParseExact2(ByVal s As String, ByVal pFormat As String, ByVal provider As IFormatProvider, ByVal style As DateTimeStyles) As DotNetLib.IDateTime
Attribute ParseExact2.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly or an exception is thrown."
    Set ParseExact2 = this.DateTimeSingleton.ParseExact2(s, pFormat, provider, style)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match at least one of the specified formats exactly or an exception is thrown.")
Public Function ParseExact3(ByVal s As String, ByRef formats() As String, ByVal provider As IFormatProvider, ByVal style As DateTimeStyles) As DotNetLib.IDateTime
Attribute ParseExact3.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match at least one of the specified formats exactly or an exception is thrown."
    Set ParseExact3 = this.DateTimeSingleton.ParseExact3(s, formats, provider, style)
End Function

'@Static
'@Description("Creates a new DateTime object that has the same number of ticks as the specified DateTime, but is designated as either local time, Coordinated Universal Time (UTC), or neither, as indicated by the specified DateTimeKind value.")
Public Function SpecifyKind(ByVal value As DotNetLib.DateTime, ByVal pKind As DotNetLib.DateTimeKind) As DotNetLib.IDateTime
Attribute SpecifyKind.VB_Description = "Creates a new DateTime object that has the same number of ticks as the specified DateTime, but is designated as either local time, Coordinated Universal Time (UTC), or neither, as indicated by the specified DateTimeKind value."
    Set SpecifyKind = this.DateTimeSingleton.SpecifyKind(value, pKind)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTime equivalent and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse(ByVal s As String, ByRef result As DotNetLib.DateTime) As Boolean
Attribute TryParse.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent and returns a value that indicates whether the conversion succeeded."
    TryParse = this.DateTimeSingleton.TryParse(s, result)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified culture-specific format information and formatting style, and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse2(ByVal s As String, ByVal provider As IFormatProvider, ByVal styles As DateTimeStyles, ByRef result As DotNetLib.DateTime) As Boolean
Attribute TryParse2.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified culture-specific format information and formatting style, and returns a value that indicates whether the conversion succeeded."
    TryParse2 = this.DateTimeSingleton.TryParse2(s, provider, styles, result)
End Function

'--------------------------------------------------------------------
'Operators
'--------------------------------------------------------------------

'@Static
'@Description("Adds a specified time interval to a specified date and time, yielding a new date and time.")
Public Function Addition(ByVal dt As DotNetLib.DateTime, ByVal ts As DotNetLib.TimeSpan) As DotNetLib.IDateTime
Attribute Addition.VB_Description = "Adds a specified time interval to a specified date and time, yielding a new date and time."
    Set Addition = this.DateTimeSingleton.Addition(dt, ts)
End Function

'@Static
'@Description("Determines whether two specified instances of DateTime are equal.")
Public Function Equality(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Boolean
Attribute Equality.VB_Description = "Determines whether two specified instances of DateTime are equal."
    Equality = this.DateTimeSingleton.Equality(d1, d2)
End Function

'@Static
'@Description("Determines whether one specified DateTime is later than another specified DateTime.")
Public Function GreaterThan(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Boolean
Attribute GreaterThan.VB_Description = "Determines whether one specified DateTime is later than another specified DateTime."
    GreaterThan = this.DateTimeSingleton.GreaterThan(d1, d2)
End Function

'@Static
'@Description("Determines whether one specified DateTime represents a date and time that is the same as or later than another specified DateTime.")
Public Function GreaterThanOrEqual(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Boolean
Attribute GreaterThanOrEqual.VB_Description = "Determines whether one specified DateTime represents a date and time that is the same as or later than another specified DateTime."
    GreaterThanOrEqual = this.DateTimeSingleton.GreaterThanOrEqual(d1, d2)
End Function

'@Static
'@Description("Determines whether two specified instances of DateTime are not equal.")
Public Function Inequality(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Boolean
Attribute Inequality.VB_Description = "Determines whether two specified instances of DateTime are not equal."
    Inequality = this.DateTimeSingleton.Inequality(d1, d2)
End Function

'@Static
'@Description("Determines whether one specified DateTime is earlier than another specified DateTime.")
Public Function LessThan(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Boolean
Attribute LessThan.VB_Description = "Determines whether one specified DateTime is earlier than another specified DateTime."
    LessThan = this.DateTimeSingleton.LessThan(d1, d2)
End Function

'@Static
'@Description("Determines whether one specified DateTime represents a date and time that is the same as or earlier than another specified DateTime.")
Public Function LessThanOrEqual(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Boolean
Attribute LessThanOrEqual.VB_Description = "Determines whether one specified DateTime represents a date and time that is the same as or earlier than another specified DateTime."
    LessThanOrEqual = this.DateTimeSingleton.LessThanOrEqual(d1, d2)
End Function

'@Static
'@Description("Subtracts a specified date and time from another specified date and time and returns a time interval.")
Public Function Subtraction(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As DotNetLib.ITimeSpan
Attribute Subtraction.VB_Description = "Subtracts a specified date and time from another specified date and time and returns a time interval."
    Set Subtraction = this.DateTimeSingleton.Subtraction(d1, d2)
End Function

'@Static
'@Description("Subtracts a specified time interval from a specified date and time and returns a new date and time.")
Public Function Subtraction2(ByVal dt As DotNetLib.DateTime, ByVal ts As DotNetLib.TimeSpan) As DotNetLib.IDateTime
Attribute Subtraction2.VB_Description = "Subtracts a specified time interval from a specified date and time and returns a new date and time."
    Set Subtraction2 = this.DateTimeSingleton.Subtraction2(dt, ts)
End Function
