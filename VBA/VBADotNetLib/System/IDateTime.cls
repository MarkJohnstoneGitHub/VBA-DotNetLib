VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Represents an instant in time, typically expressed as a date and time of day."
'Rubberduck annotations
'@Folder("VBADotNetLib.System")
'@ModuleDescription("Represents an instant in time, typically expressed as a date and time of day.")
'@Interface

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 29, 2023
'@LastModified July 30, 2023

Option Explicit

'@Description("Gets the date component of this instance.")
Public Property Get DateOnly() As DateTime
Attribute DateOnly.VB_Description = "Gets the date component of this instance."
End Property

'@Description("Gets the day of the month represented by this instance.")
Public Property Get Day() As Long
Attribute Day.VB_Description = "Gets the day of the month represented by this instance."
End Property

'@Description("Gets the day of the week represented by this instance.")
Public Property Get DayOfWeek() As DayOfWeek
Attribute DayOfWeek.VB_Description = "Gets the day of the week represented by this instance."
End Property

'@Description("Gets the day of the year represented by this instance.")
Public Property Get DayOfYear() As Long
Attribute DayOfYear.VB_Description = "Gets the day of the year represented by this instance."
End Property

'@Description("Gets the hour component of the date represented by this instance.")
Public Property Get Hour() As Long
Attribute Hour.VB_Description = "Gets the hour component of the date represented by this instance."
End Property

'@Description("Gets a value that indicates whether the time represented by this instance is based on local time, Coordinated Universal Time (UTC), or neither.")
Public Property Get Kind() As DateTimeKind
Attribute Kind.VB_Description = "Gets a value that indicates whether the time represented by this instance is based on local time, Coordinated Universal Time (UTC), or neither."
End Property

'@Description("Gets the milliseconds component of the date represented by this instance.")
Public Property Get Millisecond() As Long
Attribute Millisecond.VB_Description = "Gets the milliseconds component of the date represented by this instance."
End Property

'@Description("Gets the minute component of the date represented by this instance.")
Public Property Get Minute() As Long
Attribute Minute.VB_Description = "Gets the minute component of the date represented by this instance."
End Property

'@Description("Gets the month component of the date represented by this instance.")
Public Property Get Month() As Long
Attribute Month.VB_Description = "Gets the month component of the date represented by this instance."
End Property

'@Description("Gets the seconds component of the date represented by this instance.")
Public Property Get Second() As Long
Attribute Second.VB_Description = "Gets the seconds component of the date represented by this instance."
End Property

'@Description("Gets the number of ticks that represent the date and time of this instance.")
Public Property Get Ticks() As LongLong
Attribute Ticks.VB_Description = "Gets the number of ticks that represent the date and time of this instance."
End Property

'@Description("Gets the time of day for this instance.")
Public Property Get TimeOfDay() As TimeSpan
Attribute TimeOfDay.VB_Description = "Gets the time of day for this instance."
End Property

'@Description("Gets the year component of the date represented by this instance.")
Public Property Get Year() As Long
Attribute Year.VB_Description = "Gets the year component of the date represented by this instance."
End Property

'@Description("Returns a new DateTime that adds the value of the specified TimeSpan to the value of this instance.")
Public Function Add(ByVal value As TimeSpan) As DateTime
Attribute Add.VB_Description = "Returns a new DateTime that adds the value of the specified TimeSpan to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of days to the value of this instance.")
Public Function AddDays(ByVal value As Double) As DateTime
Attribute AddDays.VB_Description = "Returns a new DateTime that adds the specified number of days to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of hours to the value of this instance.")
Public Function AddHours(ByVal value As Double) As DateTime
Attribute AddHours.VB_Description = "Returns a new DateTime that adds the specified number of hours to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of milliseconds to the value of this instance.")
Public Function AddMilliseconds(ByVal value As Double) As DateTime
Attribute AddMilliseconds.VB_Description = "Returns a new DateTime that adds the specified number of milliseconds to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of minutes to the value of this instance.")
Public Function AddMinutes(ByVal value As Double) As DateTime
Attribute AddMinutes.VB_Description = "Returns a new DateTime that adds the specified number of minutes to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of months to the value of this instance.")
Public Function AddMonths(ByVal months As Long) As DateTime
Attribute AddMonths.VB_Description = "Returns a new DateTime that adds the specified number of months to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of seconds to the value of this instance.")
Public Function AddSeconds(ByVal value As Double) As DateTime
Attribute AddSeconds.VB_Description = "Returns a new DateTime that adds the specified number of seconds to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of ticks to the value of this instance.")
Public Function AddTicks(ByVal value As LongLong) As DateTime
Attribute AddTicks.VB_Description = "Returns a new DateTime that adds the specified number of ticks to the value of this instance."
End Function

'@Description("Returns a new DateTime that adds the specified number of years to the value of this instance.")
Public Function AddYears(ByVal value As Long) As DateTime
Attribute AddYears.VB_Description = "Returns a new DateTime that adds the specified number of years to the value of this instance."
End Function

'@Description("Compares the value of this instance to a specified DateTime value and returns an integer that indicates whether this instance is earlier than, the same as, or later than the specified DateTime value.")
Public Function CompareTo(ByVal value As DateTime) As Long
Attribute CompareTo.VB_Description = "Compares the value of this instance to a specified DateTime value and returns an integer that indicates whether this instance is earlier than, the same as, or later than the specified DateTime value."
End Function

'@Description("Returns a value indicating whether the value of this instance is equal to the value of the specified DateTime instance.")
Public Function Equals(ByVal value As DateTime) As Boolean
Attribute Equals.VB_Description = "Returns a value indicating whether the value of this instance is equal to the value of the specified DateTime instance."
End Function

'@Description("Converts the value of this instance to all the string representations supported by the standard date and time format specifiers.")
Public Function GetDateTimeFormats() As String()
Attribute GetDateTimeFormats.VB_Description = "Converts the value of this instance to all the string representations supported by the standard date and time format specifiers."
End Function

'@Description("Returns the hash code for this instance.")
Public Function GetHashCode() As Long
Attribute GetHashCode.VB_Description = "Returns the hash code for this instance."
End Function

'@Description("Indicates whether this instance of DateTime is within the daylight saving time range for the current time zone.")
Public Function IsDaylightSavingTime() As Boolean
Attribute IsDaylightSavingTime.VB_Description = "Indicates whether this instance of DateTime is within the daylight saving time range for the current time zone."
End Function

'@Description("Returns a new DateTime that subtracts the specified duration from the value of this instance.")
Public Function Subtract(ByVal value As TimeSpan) As DateTime
Attribute Subtract.VB_Description = "Returns a new DateTime that subtracts the specified duration from the value of this instance."
End Function

'@Description("Returns a new TimeSpan that subtracts the specified date and time from the value of this instance.")
Public Function Subtract2(ByVal value As DateTime) As TimeSpan
Attribute Subtract2.VB_Description = "Returns a new TimeSpan that subtracts the specified date and time from the value of this instance."
End Function

'@Description("Serializes the current DateTime object to a 64-bit binary value that subsequently can be used to recreate the DateTime object.")
Public Function ToBinary() As LongLong
Attribute ToBinary.VB_Description = "Serializes the current DateTime object to a 64-bit binary value that subsequently can be used to recreate the DateTime object."
End Function

'@Description("Converts the value of the current DateTime object to a Windows file time.")
Public Function ToFileTime() As LongLong
Attribute ToFileTime.VB_Description = "Converts the value of the current DateTime object to a Windows file time."
End Function

'@Description("Converts the value of the current DateTime object to a Windows file time.")
Public Function ToFileTimeUtc() As LongLong
Attribute ToFileTimeUtc.VB_Description = "Converts the value of the current DateTime object to a Windows file time."
End Function

'@Description("Converts the value of the current DateTime object to local time.")
Public Function ToLocalTime() As DateTime
Attribute ToLocalTime.VB_Description = "Converts the value of the current DateTime object to local time."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent long date string representation.")
Public Function ToLongDateString() As String
Attribute ToLongDateString.VB_Description = "Converts the value of the current DateTime object to its equivalent long date string representation."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent long time string representation.")
Public Function ToLongTimeString() As String
Attribute ToLongTimeString.VB_Description = "Converts the value of the current DateTime object to its equivalent long time string representation."
End Function

'@Description("Converts the value of this instance to the equivalent OLE Automation date.")
Public Function ToOADate() As Double
Attribute ToOADate.VB_Description = "Converts the value of this instance to the equivalent OLE Automation date."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent short date string representation.")
Public Function ToShortDateString() As String
Attribute ToShortDateString.VB_Description = "Converts the value of the current DateTime object to its equivalent short date string representation."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent short time string representation.")
Public Function ToShortTimeString() As String
Attribute ToShortTimeString.VB_Description = "Converts the value of the current DateTime object to its equivalent short time string representation."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent string representation using the formatting conventions of the current culture.")
Public Property Get ToString() As String
Attribute ToString.VB_Description = "Converts the value of the current DateTime object to its equivalent string representation using the formatting conventions of the current culture."
End Property

'@Description("Converts the value of the current DateTime object to its equivalent string representation using the specified format and the formatting conventions of the current culture.")
Public Function ToString2(ByVal Format As String) As String
Attribute ToString2.VB_Description = "Converts the value of the current DateTime object to its equivalent string representation using the specified format and the formatting conventions of the current culture."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent string representation using the specified culture-specific format information.")
Public Function ToString3(ByVal provider As IFormatProvider) As String
Attribute ToString3.VB_Description = "Converts the value of the current DateTime object to its equivalent string representation using the specified culture-specific format information."
End Function

'@Description("Converts the value of the current DateTime object to its equivalent string representation using the specified format and culture-specific format information.")
Public Function ToString4(ByVal Format As String, ByVal provider As IFormatProvider) As String
Attribute ToString4.VB_Description = "Converts the value of the current DateTime object to its equivalent string representation using the specified format and culture-specific format information."
End Function

'@Description("Converts the value of the current DateTime object to Coordinated Universal Time (UTC).")
Public Function ToUniversalTime() As DateTime
Attribute ToUniversalTime.VB_Description = "Converts the value of the current DateTime object to Coordinated Universal Time (UTC)."
End Function
