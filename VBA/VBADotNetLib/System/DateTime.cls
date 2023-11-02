VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "DateTime factory methods and static members."
'Rubberduck annotations
'@Folder "VBADotNetLib.System"
'@PredeclaredId
'@ModuleDescription("DateTime factory methods and static members.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v2.0 August 4, 2023
'@LastModified November 3, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1

'@Remarks
' Singleton class for DateTime containing constructors and static members.

Option Explicit

Const DateTimeType As String = "System.DateTime"

Private Type TDateTime
    DateTimeStatic  As DotNetLib.DateTimeSingleton
    DateTimeType    As DotNetLib.Type
End Type

Private this As TDateTime

Private Sub Class_Initialize()
    Set this.DateTimeStatic = New DotNetLib.DateTimeSingleton
    Set this.DateTimeType = Types.GetType(DateTimeType)
End Sub

Private Sub Class_Terminate()
    Set this.DateTimeStatic = Nothing
End Sub

'--------------------------------------------------------------------
'Constructors
'--------------------------------------------------------------------

'@Static
'@Description("Initializes a new instance of the DateTime structure to a specified number of ticks and to Coordinated Universal Time (UTC) or local time.")
'
'@Parameters
'   Ticks Int64
'       A date and time expressed in the number of 100-nanosecond intervals that
'       have elapsed since January 1, 0001 at 00:00:00.000 in the Gregorian calendar.
'   Kind DateTimeKind
'       One of the enumeration values that indicates whether ticks specifies a
'       local time, Coordinated Universal Time (UTC), or neither.
'
'@Returns DotNetLib.DateTime

'@Exceptions
'   ArgumentOutOfRangeException
'       ticks is less than DateTime.MinValue or greater than DateTime.MaxValue.
'
'   ArgumentException
'       kind is not one of the DateTimeKind values.
Public Function CreateFromTicks(ByVal pTicks As LongLong, Optional ByVal pKind As DotNetLib.DateTimeKind = DateTimeKind_Unspecified) As DotNetLib.DateTime
Attribute CreateFromTicks.VB_Description = "Initializes a new instance of the DateTime structure to a specified number of ticks and to Coordinated Universal Time (UTC) or local time."
    Set CreateFromTicks = this.DateTimeStatic.CreateFromTicks(pTicks, pKind)
End Function

'@Static
'@Description("Initializes a new instance of the DateTime structure to the specified year, month, and day.")
'
'@Parameters
'   Year Int32
'       The year (1 through 9999).
'   Month Int32
'       The month (1 through 12).
'   Day Int32
'       The day (1 through the number of days in month).
'
'@Returns DotNetLib.DateTime
'
'@Exceptions
'   ArgumentOutOfRangeException
'       year is less than 1 or greater than 9999.
'   -or-
'       month is less than 1 or greater than 12.
'   -or-
'       day is less than 1 or greater than the number of days in month.
Public Function CreateFromDate(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long) As DotNetLib.DateTime
Attribute CreateFromDate.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, and day."
    Set CreateFromDate = this.DateTimeStatic.CreateFromDate(pYear, pMonth, pDay)
End Function

'@Static
'@Description("Initializes a new instance of the DateTime structure to the specified year, month, and day for the specified calendar.")
'
'@Parameters
'   Year Int32
'       The year (1 through 9999).
'   Month Int32
'       The month (1 through 12).
'   Day Int32
'       The day (1 through the number of days in month).
'   calendar ICalendar
'       The calendar that is used to interpret year, month, and day.
'
'@Returns DotNetLib.DateTime
'
'@Exceptions
'   ArgumentNullException
'       calendar is null.
'
'   ArgumentOutOfRangeException
'       year is less than 1 or greater than 9999.
'   -or-
'       month is less than 1 or greater than 12.
'   -or-
'       day is less than 1 or greater than the number of days in month.

Public Function CreateFromDate2(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pCalendar As DotNetLib.Calendar) As DotNetLib.DateTime
Attribute CreateFromDate2.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, and day for the specified calendar."
    Set CreateFromDate2 = this.DateTimeStatic.CreateFromDate2(pYear, pMonth, pDay, pCalendar)
End Function

'@Static
'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond.")
'@Parameters
'   Year Int32
'       The year (1 through 9999).
'   Month Int32
'       The month (1 through 12).
'   Day Int32
'       The day (1 through the number of days in month).
'   Hour Int32
'       The hours (0 through 23).
'   Minute Int32
'       The minutes (0 through 59).
'   Second Int32
'       The seconds (0 through 59).
'   Millisecond Int32
'       The milliseconds (0 through 999).
'
'@Returns DotNetLib.DateTime
'
'@Exceptions
'   ArgumentOutOfRangeException
'       year is less than 1 or greater than 9999.
'   -or-
'       month is less than 1 or greater than 12.
'   -or-
'       day is less than 1 or greater than the number of days in month.
'   -or-
'       hour is less than 0 or greater than 23.
'   -or-
'       minute is less than 0 or greater than 59.
'   -or-
'       second is less than 0 or greater than 59.
'   -or-
'       millisecond is less than 0 or greater than 999.
Public Function CreateFromDateTime(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long, Optional ByVal pMillisecond As Long = 0) As DotNetLib.DateTime
Attribute CreateFromDateTime.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond."
    Set CreateFromDateTime = this.DateTimeStatic.CreateFromDateTime(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond)
End Function

'@Static
'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, and second for the specified calendar.")
Public Function CreateFromDateTime2(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long, ByVal pCalendar As DotNetLib.Calendar) As DotNetLib.DateTime
Attribute CreateFromDateTime2.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, and second for the specified calendar."
    Set CreateFromDateTime2 = this.DateTimeStatic.CreateFromDateTime2(pYear, pMonth, pDay, pHour, pMinute, pSecond, pCalendar)
End Function

'@Static
'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond for the specified calendar.")
Public Function CreateFromDateTime3(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long, ByVal pMillisecond As Long, ByVal pCalendar As DotNetLib.Calendar) As DotNetLib.DateTime
Attribute CreateFromDateTime3.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond for the specified calendar."
    Set CreateFromDateTime3 = this.DateTimeStatic.CreateFromDateTime3(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond, pCalendar)
End Function

'@Static
'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and Coordinated Universal Time (UTC) or local time.")
'
'@Parameters
'   Year Int32
'       The year (1 through 9999).
'   Month Int32
'       The month (1 through 12).
'   Day Int32
'       The day (1 through the number of days in month).
'   Hour Int32
'       The hours (0 through 23).
'   Minute Int32
'       The minutes (0 through 59).
'   Second Int32
'       The seconds (0 through 59).
'   Kind DateTimeKind
'       One of the enumeration values that indicates whether year, month, day,
'       hour, minute and second specify a local time, Coordinated Universal Time
'       (UTC), or neither.
'
'@Returns DotNetLib.DateTime
'
'@Exceptions
'
'   ArgumentOutOfRangeException
'       year is less than 1 or greater than 9999.
'   -or-
'       month is less than 1 or greater than 12.
'   -or-
'       day is less than 1 or greater than the number of days in month.
'   -or-
'       hour is less than 0 or greater than 23.
'   -or-
'       minute is less than 0 or greater than 59.
'   -or-
'       second is less than 0 or greater than 59.
'
'   ArgumentException
'       kind is not one of the DateTimeKind values.
Public Function CreateFromDateTimeKind(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long, ByVal pKind As DotNetLib.DateTimeKind) As DotNetLib.DateTime
Attribute CreateFromDateTimeKind.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and Coordinated Universal Time (UTC) or local time."
    Set CreateFromDateTimeKind = this.DateTimeStatic.CreateFromDateTimeKind(pYear, pMonth, pDay, pHour, pMinute, pSecond, pKind)
End Function

'@Static
'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time.")
'
'@Parameters
'   Year Int32
'       The year (1 through 9999).
'   Month Int32
'       The month (1 through 12).
'   Day Int32
'       The day (1 through the number of days in month).
'   Hour Int32
'       The hours (0 through 23).
'   Minute Int32
'       The minutes (0 through 59).
'   Second Int32
'       The seconds (0 through 59).
'   Millisecond Int32
'       The milliseconds (0 through 999).
'   Kind DateTimeKind
'       One of the enumeration values that indicates whether year, month, day,
'       hour, minute and second specify a local time, Coordinated Universal Time
'       (UTC), or neither.
'
'@Returns DotNetLib.DateTime
'
'@Exceptions
'
'   ArgumentOutOfRangeException
'       year is less than 1 or greater than 9999.
'   -or-
'       month is less than 1 or greater than 12.
'   -or-
'       day is less than 1 or greater than the number of days in month.
'   -or-
'       hour is less than 0 or greater than 23.
'   -or-
'       minute is less than 0 or greater than 59.
'   -or-
'       second is less than 0 or greater than 59.
'   -or-
'       millisecond is less than 0 or greater than 999.
'
'   ArgumentException
'       kind is not one of the DateTimeKind values.
Public Function CreateFromDateTimeKind2(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long, ByVal pMillisecond As Long, ByVal pKind As DotNetLib.DateTimeKind) As DotNetLib.DateTime
Attribute CreateFromDateTimeKind2.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time."
    Set CreateFromDateTimeKind2 = this.DateTimeStatic.CreateFromDateTimeKind2(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond, pKind)
End Function

'@Static
'@Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time for the specified calendar.")
Public Function CreateFromDateTimeKind3(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long, ByVal pMillisecond As Long, ByVal pCalendar As DotNetLib.Calendar, ByVal pKind As DotNetLib.DateTimeKind) As DotNetLib.DateTime
Attribute CreateFromDateTimeKind3.VB_Description = "Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time for the specified calendar."
    Set CreateFromDateTimeKind3 = this.DateTimeStatic.CreateFromDateTimeKind3(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond, pCalendar, pKind)
End Function

'--------------------------------------------------------------------
'Fields
'--------------------------------------------------------------------

'@Static
'@Description("Represents the largest possible value of DateTime. This field is read-only.")
Public Property Get MaxValue() As DotNetLib.DateTime
Attribute MaxValue.VB_Description = "Represents the largest possible value of DateTime. This field is read-only."
    Set MaxValue = this.DateTimeStatic.MaxValue
End Property

'@Static
'@Description("Represents the smallest possible value of DateTime. This field is read-only.")
Public Property Get MinValue() As DotNetLib.DateTime
Attribute MinValue.VB_Description = "Represents the smallest possible value of DateTime. This field is read-only."
    Set MinValue = this.DateTimeStatic.MinValue
End Property

'--------------------------------------------------------------------
'Properties
'--------------------------------------------------------------------
'@Static
'@Description("Gets a DateTime object that is set to the current date and time on this computer, expressed as the local time.")
Public Property Get Now() As DotNetLib.DateTime
Attribute Now.VB_Description = "Gets a DateTime object that is set to the current date and time on this computer, expressed as the local time."
    Set Now = this.DateTimeStatic.Now()
End Property

'@Static
'@Description("Gets the current date.")
Public Property Get Today() As DotNetLib.DateTime
Attribute Today.VB_Description = "Gets the current date."
    Set Today = this.DateTimeStatic.Today()
End Property

'@Static
'@Description("Gets a DateTime object that is set to the current date and time on this computer, expressed as the Coordinated Universal Time (UTC).")
Public Property Get UtcNow() As DotNetLib.DateTime
Attribute UtcNow.VB_Description = "Gets a DateTime object that is set to the current date and time on this computer, expressed as the Coordinated Universal Time (UTC)."
    Set UtcNow = this.DateTimeStatic.UtcNow()
End Property

'--------------------------------------------------------------------
'Methods
'--------------------------------------------------------------------

'@Static
'@Description("Compares two instances of DateTime and returns an integer that indicates whether the first instance is earlier than, the same as, or later than the second instance.")
Public Function Compare(ByVal t1 As DotNetLib.DateTime, ByVal t2 As DotNetLib.DateTime) As Long
Attribute Compare.VB_Description = "Compares two instances of DateTime and returns an integer that indicates whether the first instance is earlier than, the same as, or later than the second instance."
    Compare = this.DateTimeStatic.Compare(t1, t2)
End Function

'@Static
'@Description("Returns the number of days in the specified month and year.")
Public Function DaysInMonth(ByVal pYear As Long, ByVal pMonth As Long) As Long
Attribute DaysInMonth.VB_Description = "Returns the number of days in the specified month and year."
    DaysInMonth = this.DateTimeStatic.DaysInMonth(pYear, pMonth)
End Function

'@Static
'@Description("Returns a value indicating whether two DateTime instances have the same date and time value.")
Public Function Equals(ByVal t1 As DotNetLib.DateTime, ByVal t2 As DotNetLib.DateTime) As Boolean
Attribute Equals.VB_Description = "Returns a value indicating whether two DateTime instances have the same date and time value."
    Equals = this.DateTimeStatic.Equals(t1, t2)
End Function

'@Static
'@Description("Deserializes a 64-bit binary value and recreates an original serialized DateTime object.")
Public Function FromBinary(ByVal dateData As LongLong) As DotNetLib.DateTime
Attribute FromBinary.VB_Description = "Deserializes a 64-bit binary value and recreates an original serialized DateTime object."
    Set FromBinary = this.DateTimeStatic.FromBinary(dateData)
End Function

'@Static
'@Description("Converts the specified Windows file time to an equivalent local time.")
Public Function FromFileTime(ByVal fileTime As LongLong) As DotNetLib.DateTime
Attribute FromFileTime.VB_Description = "Converts the specified Windows file time to an equivalent local time."
    Set FromFileTime = this.DateTimeStatic.FromFileTime(fileTime)
End Function

'@Static
'@Description("Converts the specified Windows file time to an equivalent UTC time.")
Public Function FromFileTimeUtc(ByVal fileTime As LongLong) As DotNetLib.DateTime
Attribute FromFileTimeUtc.VB_Description = "Converts the specified Windows file time to an equivalent UTC time."
    Set FromFileTimeUtc = this.DateTimeStatic.FromFileTimeUtc(fileTime)
End Function

'@Static
'@Description("Returns a DateTime equivalent to the specified OLE Automation Date.")
Public Function FromOADate(ByVal oaDate As Date) As DotNetLib.DateTime
Attribute FromOADate.VB_Description = "Returns a DateTime equivalent to the specified OLE Automation Date."
    Set FromOADate = this.DateTimeStatic.FromOADate(oaDate)
End Function

'@Static
'@Description("Returns an indication whether the specified year is a leap year.")
Public Function IsLeapYear(ByVal pYear As Long) As Boolean
Attribute IsLeapYear.VB_Description = "Returns an indication whether the specified year is a leap year."
    IsLeapYear = this.DateTimeStatic.IsLeapYear(pYear)
End Function

'@Static
'@Description("Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture.")
Public Function Parse(ByVal s As String) As DotNetLib.DateTime
Attribute Parse.VB_Description = "Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture."
    Set Parse = this.DateTimeStatic.Parse(s)
End Function

'@Static
'@Description("Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture.")
Public Function Parse2(ByVal s As String, ByVal provider As IFormatProvider) As DotNetLib.DateTime
Attribute Parse2.VB_Description = "Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture."
    Set Parse2 = this.DateTimeStatic.Parse2(s, provider)
End Function

'@Static
'@Description("Defines the formatting options that customize string parsing for some date and time parsing methods.  This enumeration supports a bitwise combination of its member values.")
Public Function Parse3(ByVal s As String, ByVal provider As IFormatProvider, ByVal styles As DateTimeStyles) As DotNetLib.DateTime
Attribute Parse3.VB_Description = "Defines the formatting options that customize string parsing for some date and time parsing methods.  This enumeration supports a bitwise combination of its member values."
    Set Parse3 = this.DateTimeStatic.Parse3(s, provider, styles)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")
Public Function ParseExact(ByVal s As String, ByVal pFormat As String, ByVal provider As IFormatProvider) As DotNetLib.DateTime
Attribute ParseExact.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly."
    Set ParseExact = this.DateTimeStatic.ParseExact(s, pFormat, provider)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly or an exception is thrown.")
Public Function ParseExact2(ByVal s As String, ByVal pFormat As String, ByVal provider As IFormatProvider, ByVal style As DateTimeStyles) As DotNetLib.DateTime
Attribute ParseExact2.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly or an exception is thrown."
    Set ParseExact2 = this.DateTimeStatic.ParseExact2(s, pFormat, provider, style)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match at least one of the specified formats exactly or an exception is thrown.")
Public Function ParseExact3(ByVal s As String, ByRef formats() As String, ByVal provider As IFormatProvider, ByVal style As DateTimeStyles) As DotNetLib.IDateTime
Attribute ParseExact3.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match at least one of the specified formats exactly or an exception is thrown."
    Set ParseExact3 = this.DateTimeStatic.ParseExact3(s, formats, provider, style)
End Function

'@Static
'@Description("Creates a new DateTime object that has the same number of ticks as the specified DateTime, but is designated as either local time, Coordinated Universal Time (UTC), or neither, as indicated by the specified DateTimeKind value.")
Public Function SpecifyKind(ByVal value As DotNetLib.DateTime, ByVal pKind As DotNetLib.DateTimeKind) As DotNetLib.DateTime
Attribute SpecifyKind.VB_Description = "Creates a new DateTime object that has the same number of ticks as the specified DateTime, but is designated as either local time, Coordinated Universal Time (UTC), or neither, as indicated by the specified DateTimeKind value."
    Set SpecifyKind = this.DateTimeStatic.SpecifyKind(value, pKind)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTime equivalent and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse(ByVal s As String, ByRef result As DotNetLib.DateTime) As Boolean
Attribute TryParse.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent and returns a value that indicates whether the conversion succeeded."
    TryParse = this.DateTimeStatic.TryParse(s, result)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified culture-specific format information and formatting style, and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse2(ByVal s As String, ByVal provider As IFormatProvider, ByVal styles As DateTimeStyles, ByRef result As DotNetLib.DateTime) As Boolean
Attribute TryParse2.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified culture-specific format information and formatting style, and returns a value that indicates whether the conversion succeeded."
    TryParse2 = this.DateTimeStatic.TryParse2(s, provider, styles, result)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified format, culture-specific format information, and style.")
Public Function TryParseExact(ByVal s As String, ByVal pFormat As String, ByVal provider As IFormatProvider, ByVal style As DateTimeStyles, ByRef result As DotNetLib.DateTime) As Boolean
Attribute TryParseExact.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified format, culture-specific format information, and style."
    TryParseExact = this.DateTimeStatic.TryParseExact(s, pFormat, provider, style, result)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified array of formats, culture-specific format information, and style.")
Public Function TryParseExact2(ByVal s As String, ByRef formats() As String, ByVal provider As IFormatProvider, ByVal style As DateTimeStyles, ByRef result As DotNetLib.DateTime) As Boolean
Attribute TryParseExact2.VB_Description = "Converts the specified string representation of a date and time to its DateTime equivalent using the specified array of formats, culture-specific format information, and style."
    TryParseExact2 = this.DateTimeStatic.TryParseExact2(s, formats, provider, style, result)
End Function

'--------------------------------------------------------------------
'Operators
'--------------------------------------------------------------------

'@Static
'@Description("Adds a specified time interval to a specified date and time, yielding a new date and time.")
Public Function Addition(ByVal dt As DotNetLib.DateTime, ByVal ts As DotNetLib.TimeSpan) As DotNetLib.DateTime
Attribute Addition.VB_Description = "Adds a specified time interval to a specified date and time, yielding a new date and time."
    Set Addition = this.DateTimeStatic.Addition(dt, ts)
End Function

'@Static
'@Description("Determines whether two specified instances of DateTime are equal.")
Public Function Equality(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Boolean
Attribute Equality.VB_Description = "Determines whether two specified instances of DateTime are equal."
    Equality = this.DateTimeStatic.Equality(d1, d2)
End Function

'@Static
'@Description("Determines whether one specified DateTime is later than another specified DateTime.")
Public Function GreaterThan(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Boolean
Attribute GreaterThan.VB_Description = "Determines whether one specified DateTime is later than another specified DateTime."
    GreaterThan = this.DateTimeStatic.GreaterThan(d1, d2)
End Function

'@Static
'@Description("Determines whether one specified DateTime represents a date and time that is the same as or later than another specified DateTime.")
Public Function GreaterThanOrEqual(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Boolean
Attribute GreaterThanOrEqual.VB_Description = "Determines whether one specified DateTime represents a date and time that is the same as or later than another specified DateTime."
    GreaterThanOrEqual = this.DateTimeStatic.GreaterThanOrEqual(d1, d2)
End Function

'@Static
'@Description("Determines whether two specified instances of DateTime are not equal.")
Public Function Inequality(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Boolean
Attribute Inequality.VB_Description = "Determines whether two specified instances of DateTime are not equal."
    Inequality = this.DateTimeStatic.Inequality(d1, d2)
End Function

'@Static
'@Description("Determines whether one specified DateTime is earlier than another specified DateTime.")
Public Function LessThan(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Boolean
Attribute LessThan.VB_Description = "Determines whether one specified DateTime is earlier than another specified DateTime."
    LessThan = this.DateTimeStatic.LessThan(d1, d2)
End Function

'@Static
'@Description("Determines whether one specified DateTime represents a date and time that is the same as or earlier than another specified DateTime.")
Public Function LessThanOrEqual(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Boolean
Attribute LessThanOrEqual.VB_Description = "Determines whether one specified DateTime represents a date and time that is the same as or earlier than another specified DateTime."
    LessThanOrEqual = this.DateTimeStatic.LessThanOrEqual(d1, d2)
End Function

'@Static
'@Description("Subtracts a specified date and time from another specified date and time and returns a time interval.")
Public Function Subtraction(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As DotNetLib.TimeSpan
Attribute Subtraction.VB_Description = "Subtracts a specified date and time from another specified date and time and returns a time interval."
    Set Subtraction = this.DateTimeStatic.Subtraction(d1, d2)
End Function

'@Static
'@Description("Subtracts a specified time interval from a specified date and time and returns a new date and time.")
Public Function Subtraction2(ByVal dt As DotNetLib.DateTime, ByVal ts As DotNetLib.TimeSpan) As DotNetLib.DateTime
Attribute Subtraction2.VB_Description = "Subtracts a specified time interval from a specified date and time and returns a new date and time."
    Set Subtraction2 = this.DateTimeStatic.Subtraction2(dt, ts)
End Function

'--------------------------------------------------------------------
'Extension
'Added to get the type for arrays etc.
'--------------------------------------------------------------------
'@Static
'@DefaultMember
'@Description("Gets the DateTime Type.")
Public Function GetType() As DotNetLib.Type
Attribute GetType.VB_Description = "Gets the DateTime Type."
Attribute GetType.VB_UserMemId = 0
    Set GetType = this.DateTimeType
End Function

'@Static
'@Description("Gets the total years difference between two DateTime objects.")
'@Reference
' https://stackoverflow.com/a/1595311/10759363
' https://stackoverflow.com/questions/9/how-do-i-calculate-someones-age-based-on-a-datetime-type-birthday
Public Function TotalYears(ByVal d1 As DotNetLib.DateTime, ByVal d2 As DotNetLib.DateTime) As Long
Attribute TotalYears.VB_Description = "Gets the total years difference between two DateTime objects."
    Dim result As Long
    If d1.Ticks > d2.Ticks Then
        result = d1.Year - d2.Year
        If (d2.Ticks > d1.AddYears(-result).Ticks) Then
            result = result - 1
        End If
        result = result * -1
    Else
        result = d2.Year - d1.Year
        If (d1.Ticks > d2.AddYears(-result).Ticks) Then
            result = result - 1
        End If
    End If
    TotalYears = result
End Function
