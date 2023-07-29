VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDateTimeOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a point in time, typically expressed as a date and time of day, relative to Coordinated Universal Time (UTC)."
'@Folder("VBADotNetLib.System")
'@ModuleDescription("Represents a point in time, typically expressed as a date and time of day, relative to Coordinated Universal Time (UTC).")
'@Interface

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 29, 2023
'@LastModified July 29, 2023

Option Explicit

'--------------------------------------------------------------------
'Properties
'--------------------------------------------------------------------

'@Description("Gets a DateTime value that represents the date component of the current DateTimeOffset object.")
Public Property Get DateOnly() As IDateTime
Attribute DateOnly.VB_Description = "Gets a DateTime value that represents the date component of the current DateTimeOffset object."
End Property

'@Description("Gets a DateTime value that represents the date and time of the current DateTimeOffset object.")
Public Property Get DateTime() As IDateTime
Attribute DateTime.VB_Description = "Gets a DateTime value that represents the date and time of the current DateTimeOffset object."
End Property

'@Description("Gets the day of the month represented by the current DateTimeOffset object.")
Public Property Get Day() As Long
Attribute Day.VB_Description = "Gets the day of the month represented by the current DateTimeOffset object."
End Property

'@Description("Gets the day of the week represented by the current DateTimeOffset object.")
Public Property Get DayOfWeek() As DayOfWeek
Attribute DayOfWeek.VB_Description = "Gets the day of the week represented by the current DateTimeOffset object."
End Property

'@Description("Gets the day of the year represented by the current DateTimeOffset object.")
Public Property Get DayOfYear() As Long
Attribute DayOfYear.VB_Description = "Gets the day of the year represented by the current DateTimeOffset object."
End Property

'@Description("Gets the hour component of the time represented by the current DateTimeOffset object.")
Public Property Get Hour() As Long
Attribute Hour.VB_Description = "Gets the hour component of the time represented by the current DateTimeOffset object."
End Property

'@Description("Gets a DateTime value that represents the local date and time of the current DateTimeOffset object.")
Public Property Get LocalDateTime() As IDateTime
Attribute LocalDateTime.VB_Description = "Gets a DateTime value that represents the local date and time of the current DateTimeOffset object."
End Property

'@Description("Gets the millisecond component of the time represented by the current DateTimeOffset object.")
Public Property Get Millisecond() As Long
Attribute Millisecond.VB_Description = "Gets the millisecond component of the time represented by the current DateTimeOffset object."
End Property

'@Description("Gets the minute component of the time represented by the current DateTimeOffset object.")
Public Property Get Minute() As Long
Attribute Minute.VB_Description = "Gets the minute component of the time represented by the current DateTimeOffset object."
End Property

'@Description("Gets the month component of the date represented by the current DateTimeOffset object.")
Public Property Get Month() As Long
Attribute Month.VB_Description = "Gets the month component of the date represented by the current DateTimeOffset object."
End Property

'@Description("Gets the time's offset from Coordinated Universal Time (UTC).")
Public Property Get Offset() As ITimeSpan
Attribute Offset.VB_Description = "Gets the time's offset from Coordinated Universal Time (UTC)."
End Property

'@Description("Gets the second component of the clock time represented by the current DateTimeOffset object.")
Public Property Get Second() As Long
Attribute Second.VB_Description = "Gets the second component of the clock time represented by the current DateTimeOffset object."
End Property

'@Description("Gets the number of ticks that represents the date and time of the current DateTimeOffset object in clock time.")
Public Property Get Ticks() As LongLong
Attribute Ticks.VB_Description = "Gets the number of ticks that represents the date and time of the current DateTimeOffset object in clock time."
End Property

'@Description("Gets the time of day for the current DateTimeOffset object.")
Public Property Get TimeOfDay() As ITimeSpan
Attribute TimeOfDay.VB_Description = "Gets the time of day for the current DateTimeOffset object."
End Property

'@Description("Gets a DateTime value that represents the Coordinated Universal Time (UTC) date and time of the current DateTimeOffset object.")
Public Property Get UtcDateTime() As IDateTime
Attribute UtcDateTime.VB_Description = "Gets a DateTime value that represents the Coordinated Universal Time (UTC) date and time of the current DateTimeOffset object."
End Property

'@Description("Gets the number of ticks that represents the date and time of the current DateTimeOffset object in Coordinated Universal Time (UTC).")
Public Property Get UtcTicks() As LongLong
Attribute UtcTicks.VB_Description = "Gets the number of ticks that represents the date and time of the current DateTimeOffset object in Coordinated Universal Time (UTC)."
End Property

'@Description("Gets the year component of the date represented by the current DateTimeOffset object.")
Public Property Get Year() As Long
Attribute Year.VB_Description = "Gets the year component of the date represented by the current DateTimeOffset object."
End Property

'--------------------------------------------------------------------
'Methods
'--------------------------------------------------------------------
'@Description("Returns a new DateTimeOffset object that adds a specified time interval to the value of this instance.")
Public Function Add(ByVal TimeSpan As TimeSpan) As IDateTimeOffset
Attribute Add.VB_Description = "Returns a new DateTimeOffset object that adds a specified time interval to the value of this instance."
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional days to the value of this instance.")
Public Function AddDays(ByVal Days As Double) As IDateTimeOffset
Attribute AddDays.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of whole and fractional days to the value of this instance."
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional hours to the value of this instance.")
Public Function AddHours(ByVal Hours As Double) As IDateTimeOffset
Attribute AddHours.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of whole and fractional hours to the value of this instance."
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of milliseconds to the value of this instance.")
Public Function AddMilliseconds(ByVal Milliseconds As Double) As IDateTimeOffset
Attribute AddMilliseconds.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of milliseconds to the value of this instance."
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional minutes to the value of this instance.")
Public Function AddMinutes(ByVal Minutes As Double) As IDateTimeOffset
Attribute AddMinutes.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of whole and fractional minutes to the value of this instance."
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of months to the value of this instance.")
Public Function AddMonths(ByVal months As Long) As IDateTimeOffset
Attribute AddMonths.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of months to the value of this instance."
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional seconds to the value of this instance.")
Public Function AddSeconds(ByVal Seconds As Double) As IDateTimeOffset
Attribute AddSeconds.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of whole and fractional seconds to the value of this instance."
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of ticks to the value of this instance.")
Public Function AddTicks(ByVal Ticks As LongLong) As IDateTimeOffset
Attribute AddTicks.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of ticks to the value of this instance."
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of years to the value of this instance.")
Public Function AddYears(ByVal years As Long) As IDateTimeOffset
Attribute AddYears.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of years to the value of this instance."
End Function

'@Description("Compares the current DateTimeOffset object to a specified DateTimeOffset object and indicates whether the current object is earlier than, the same as, or later than the second DateTimeOffset object.")
Public Function CompareTo(ByVal other As DateTimeOffset) As Long
Attribute CompareTo.VB_Description = "Compares the current DateTimeOffset object to a specified DateTimeOffset object and indicates whether the current object is earlier than, the same as, or later than the second DateTimeOffset object."
End Function

'@Description("Determines whether the current DateTimeOffset object represents the same point in time as a specified DateTimeOffset object.")
Public Function Equals(ByVal other As DateTimeOffset) As Boolean
Attribute Equals.VB_Description = "Determines whether the current DateTimeOffset object represents the same point in time as a specified DateTimeOffset object."
End Function

'@Description("Determines whether the current DateTimeOffset object represents the same time and has the same offset as a specified DateTimeOffset object.")
Public Function EqualsExact(ByVal other As DateTimeOffset) As Boolean
Attribute EqualsExact.VB_Description = "Determines whether the current DateTimeOffset object represents the same time and has the same offset as a specified DateTimeOffset object."
End Function

'@Description("Returns the hash code for the current DateTimeOffset object.")
Public Function GetHashCode() As Long
Attribute GetHashCode.VB_Description = "Returns the hash code for the current DateTimeOffset object."
End Function

'@Description("Subtracts a DateTimeOffset value that represents a specific date and time from the current DateTimeOffset object.")
Public Function Subtract(ByVal value As DateTimeOffset) As ITimeSpan
Attribute Subtract.VB_Description = "Subtracts a DateTimeOffset value that represents a specific date and time from the current DateTimeOffset object."
End Function

'@Description("Subtracts a specified time interval from the current DateTimeOffset object.")
Public Function Subtract2(ByVal value As TimeSpan) As IDateTimeOffset
Attribute Subtract2.VB_Description = "Subtracts a specified time interval from the current DateTimeOffset object."
End Function

'@Description("Converts the value of the current DateTimeOffset object to a Windows file time.")
Public Function ToFileTime() As LongLong
Attribute ToFileTime.VB_Description = "Converts the value of the current DateTimeOffset object to a Windows file time."
End Function

'@Description("Converts the current DateTimeOffset object to a DateTimeOffset object that represents the local time.")
Public Function ToLocalTime() As IDateTimeOffset
Attribute ToLocalTime.VB_Description = "Converts the current DateTimeOffset object to a DateTimeOffset object that represents the local time."
End Function

'@Description("Converts the value of the current DateTimeOffset object to the date and time specified by an offset value.")
Public Function ToOffset(ByVal Offset As TimeSpan) As IDateTimeOffset
Attribute ToOffset.VB_Description = "Converts the value of the current DateTimeOffset object to the date and time specified by an offset value."
End Function

'@DefaultMember
'@Description("Converts the value of the current DateTimeOffset object to its equivalent string representation.")
Public Property Get ToString() As String
Attribute ToString.VB_Description = "Converts the value of the current DateTimeOffset object to its equivalent string representation."
Attribute ToString.VB_UserMemId = 0
End Property

'@Description("Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified format.")
Public Function ToString2(ByVal Format As String) As String
Attribute ToString2.VB_Description = "Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified format."
End Function

'@Description("Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified culture-specific formatting information.")
Public Function ToString3(ByVal formatProvider As IFormatProvider) As String
Attribute ToString3.VB_Description = "Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified culture-specific formatting information."
End Function

'@Description("Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified format and culture-specific format information.")
Public Function ToString4(ByVal Format As String, ByVal formatProvider As IFormatProvider) As String
Attribute ToString4.VB_Description = "Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified format and culture-specific format information."
End Function

'@Description("Converts the current DateTimeOffset object to a DateTimeOffset value that represents the Coordinated Universal Time (UTC).")
Public Function ToUniversalTime() As DateTimeOffset
Attribute ToUniversalTime.VB_Description = "Converts the current DateTimeOffset object to a DateTimeOffset value that represents the Coordinated Universal Time (UTC)."
End Function

'@Description("Returns the number of milliseconds that have elapsed since 1970-01-01T00:00:00.000Z.")
Public Function ToUnixTimeMilliseconds() As LongLong
Attribute ToUnixTimeMilliseconds.VB_Description = "Returns the number of milliseconds that have elapsed since 1970-01-01T00:00:00.000Z."
End Function

'@Description("Returns the number of seconds that have elapsed since 1970-01-01T00:00:00Z.")
Public Function ToUnixTimeSeconds() As LongLong
Attribute ToUnixTimeSeconds.VB_Description = "Returns the number of seconds that have elapsed since 1970-01-01T00:00:00Z."
End Function
