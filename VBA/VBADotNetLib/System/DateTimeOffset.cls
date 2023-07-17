VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a point in time, typically expressed as a date and time of day, relative to Coordinated Universal Time (UTC)."
'@Folder "VBADotNetLib.System"
'@ModuleDescription("Represents a point in time, typically expressed as a date and time of day, relative to Coordinated Universal Time (UTC).")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 11, 2023
'@LastModified July 17, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1

'@Remarks
' Generated from DotNetLib.tlb using https://github.com/MarkJohnstoneGitHub/Refactor-COM-object-to-VBA-COM-wrapper-class

Option Explicit

Private Type TDateTimeOffset
   DotNetLibDateTimeOffset As DotNetLib.DateTimeOffset
End Type

Private this As TDateTimeOffset

Private Sub Class_Initialize()
   Set this.DotNetLibDateTimeOffset = New DotNetLib.DateTimeOffset
End Sub

Private Sub Class_Terminate()
   Set this.DotNetLibDateTimeOffset = Nothing
End Sub

Friend Property Get ComObject() As DotNetLib.DateTimeOffset
   Set ComObject = this.DotNetLibDateTimeOffset
End Property

Friend Property Set ComObject(ByVal objDateTimeOffset As DotNetLib.DateTimeOffset)
   Set this.DotNetLibDateTimeOffset = objDateTimeOffset
End Property

Friend Property Get Self() As DateTimeOffset
   Set Self = Me
End Property

'@Description("Initializes a new instance of the DateTimeOffset structure using the specified DateTime value.")
Public Function CreateFromDateTime(ByVal DateTime As DateTime) As DateTimeOffset
Attribute CreateFromDateTime.VB_Description = "Initializes a new instance of the DateTimeOffset structure using the specified DateTime value."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.CreateFromDateTime(DateTime.ComObject)
      Set CreateFromDateTime = .Self
   End With
End Function

'@Description("Initializes a new instance of the DateTimeOffset structure using the specified DateTime value and offset.")
'@Ignore UseMeaningfulName
Public Function CreateFromDateTime2(ByVal DateTime As DateTime, ByVal Offset As TimeSpan) As DateTimeOffset
Attribute CreateFromDateTime2.VB_Description = "Initializes a new instance of the DateTimeOffset structure using the specified DateTime value and offset."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.CreateFromDateTime2(DateTime.ComObject, Offset.ComObject)
      Set CreateFromDateTime2 = .Self
   End With
End Function

'@Description("Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, and offset.")
Public Function CreateFromDateTimeParts(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Offset As TimeSpan) As DateTimeOffset
Attribute CreateFromDateTimeParts.VB_Description = "Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, and offset."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.CreateFromDateTimeParts(Year, Month, Day, Hour, Minute, Second, Offset.ComObject)
      Set CreateFromDateTimeParts = .Self
   End With
End Function

'@Description("Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, millisecond, and offset.")
'@Ignore UseMeaningfulName
Public Function CreateFromDateTimeParts2(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, ByVal Offset As TimeSpan) As DateTimeOffset
Attribute CreateFromDateTimeParts2.VB_Description = "Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, millisecond, and offset."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.CreateFromDateTimeParts2(Year, Month, Day, Hour, Minute, Second, Millisecond, Offset.ComObject)
      Set CreateFromDateTimeParts2 = .Self
   End With
End Function

'@Ignore UseMeaningfulName
Public Function CreateFromDateTimeParts3(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, ByVal calendar As calendar, ByVal Offset As TimeSpan) As DateTimeOffset
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.CreateFromDateTimeParts3(Year, Month, Day, Hour, Minute, Second, Millisecond, calendar, Offset.ComObject)
      Set CreateFromDateTimeParts3 = .Self
   End With
End Function

'@Description("Initializes a new instance of the DateTimeOffset structure using the specified number of ticks and offset.")
Public Function CreateFromTicks(ByVal Ticks As LongLong, ByVal Offset As TimeSpan) As DateTimeOffset
Attribute CreateFromTicks.VB_Description = "Initializes a new instance of the DateTimeOffset structure using the specified number of ticks and offset."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.CreateFromTicks(Ticks, Offset.ComObject)
      Set CreateFromTicks = .Self
   End With
End Function

'@Description("Represents the greatest possible value of DateTimeOffset. This field is read-only.")
Public Property Get MaxValue() As DateTimeOffset
Attribute MaxValue.VB_Description = "Represents the greatest possible value of DateTimeOffset. This field is read-only."
   Set MaxValue = DateTimeOffsetStatic.MaxValue
End Property

'@Description("Represents the earliest possible DateTimeOffset value. This field is read-only.")
Public Property Get MinValue() As DateTimeOffset
Attribute MinValue.VB_Description = "Represents the earliest possible DateTimeOffset value. This field is read-only."
   Set MinValue = DateTimeOffsetStatic.MinValue
End Property

'@Description("Gets a DateTime value that represents the date component of the current DateTimeOffset object.")
Public Property Get DateOnly() As DateTime
Attribute DateOnly.VB_Description = "Gets a DateTime value that represents the date component of the current DateTimeOffset object."
   With New DateTime
      Set .ComObject = this.DotNetLibDateTimeOffset.Date()
      Set DateOnly = .Self
   End With
End Property

'@Description("Gets a DateTime value that represents the date and time of the current DateTimeOffset object.")
Public Property Get DateTime() As DateTime
Attribute DateTime.VB_Description = "Gets a DateTime value that represents the date and time of the current DateTimeOffset object."
   With New DateTime
      Set .ComObject = this.DotNetLibDateTimeOffset.Date()
      Set DateTime = .Self
   End With
End Property

'@Description("Gets the day of the month represented by the current DateTimeOffset object.")
Public Property Get Day() As Long
Attribute Day.VB_Description = "Gets the day of the month represented by the current DateTimeOffset object."
   Day = this.DotNetLibDateTimeOffset.Day()
End Property

'@Description("Gets the day of the week represented by the current DateTimeOffset object.")
Public Property Get DayOfWeek() As DotNetLib.DayOfWeek
Attribute DayOfWeek.VB_Description = "Gets the day of the week represented by the current DateTimeOffset object."
   DayOfWeek = this.DotNetLibDateTimeOffset.DayOfWeek()
End Property

'@Description("Gets the day of the year represented by the current DateTimeOffset object.")
Public Property Get DayOfYear() As Long
Attribute DayOfYear.VB_Description = "Gets the day of the year represented by the current DateTimeOffset object."
   DayOfYear = this.DotNetLibDateTimeOffset.DayOfYear()
End Property

'@Description("Gets the hour component of the time represented by the current DateTimeOffset object.")
Public Property Get Hour() As Long
Attribute Hour.VB_Description = "Gets the hour component of the time represented by the current DateTimeOffset object."
   Hour = this.DotNetLibDateTimeOffset.Hour()
End Property

'@Description("Gets a DateTime value that represents the local date and time of the current DateTimeOffset object.")
Public Property Get localDateTime() As DateTime
Attribute localDateTime.VB_Description = "Gets a DateTime value that represents the local date and time of the current DateTimeOffset object."
   With New DateTime
      Set .ComObject = this.DotNetLibDateTimeOffset.localDateTime()
      Set localDateTime = .Self
   End With
End Property

'@Description("Gets the millisecond component of the time represented by the current DateTimeOffset object.")
Public Property Get Millisecond() As Long
Attribute Millisecond.VB_Description = "Gets the millisecond component of the time represented by the current DateTimeOffset object."
   Millisecond = this.DotNetLibDateTimeOffset.Millisecond()
End Property

'@Description("Gets the minute component of the time represented by the current DateTimeOffset object.")
Public Property Get Minute() As Long
Attribute Minute.VB_Description = "Gets the minute component of the time represented by the current DateTimeOffset object."
   Minute = this.DotNetLibDateTimeOffset.Minute()
End Property

'@Description("Gets the month component of the date represented by the current DateTimeOffset object.")
Public Property Get Month() As Long
Attribute Month.VB_Description = "Gets the month component of the date represented by the current DateTimeOffset object."
   Month = this.DotNetLibDateTimeOffset.Month()
End Property

'@Description("Gets a DateTimeOffset object that is set to the current date and time on the current computer, with the offset set to the local time's offset from Coordinated Universal Time (UTC).")
Public Property Get Now() As DateTimeOffset
Attribute Now.VB_Description = "Gets a DateTimeOffset object that is set to the current date and time on the current computer, with the offset set to the local time's offset from Coordinated Universal Time (UTC)."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.Now()
      Set Now = .Self
   End With
End Property

'@Description("Gets the time's offset from Coordinated Universal Time (UTC).")
Public Property Get Offset() As TimeSpan
Attribute Offset.VB_Description = "Gets the time's offset from Coordinated Universal Time (UTC)."
   With New TimeSpan
      Set .ComObject = this.DotNetLibDateTimeOffset.Offset()
      Set Offset = .Self
   End With
End Property

'@Description("Gets the second component of the clock time represented by the current DateTimeOffset object.")
Public Property Get Second() As Long
Attribute Second.VB_Description = "Gets the second component of the clock time represented by the current DateTimeOffset object."
   Second = this.DotNetLibDateTimeOffset.Second()
End Property

'@Description("Gets the number of ticks that represents the date and time of the current DateTimeOffset object in clock time.")
Public Property Get Ticks() As LongLong
Attribute Ticks.VB_Description = "Gets the number of ticks that represents the date and time of the current DateTimeOffset object in clock time."
   Ticks = this.DotNetLibDateTimeOffset.Ticks()
End Property

'@Description("Gets the time of day for the current DateTimeOffset object.")
Public Property Get TimeOfDay() As TimeSpan
Attribute TimeOfDay.VB_Description = "Gets the time of day for the current DateTimeOffset object."
   With New TimeSpan
      Set .ComObject = this.DotNetLibDateTimeOffset.TimeOfDay()
      Set TimeOfDay = .Self
   End With
End Property

'@Description("Gets a DateTime value that represents the Coordinated Universal Time (UTC) date and time of the current DateTimeOffset object.")
Public Property Get UtcDateTime() As DateTime
Attribute UtcDateTime.VB_Description = "Gets a DateTime value that represents the Coordinated Universal Time (UTC) date and time of the current DateTimeOffset object."
   With New DateTime
      Set .ComObject = this.DotNetLibDateTimeOffset.UtcDateTime()
      Set UtcDateTime = .Self
   End With
End Property

'@Description("Gets a DateTimeOffset object whose date and time are set to the current Coordinated Universal Time (UTC) date and time and whose offset is Zero.")
Public Property Get UtcNow() As DateTimeOffset
Attribute UtcNow.VB_Description = "Gets a DateTimeOffset object whose date and time are set to the current Coordinated Universal Time (UTC) date and time and whose offset is Zero."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.UtcNow()
      Set UtcNow = .Self
   End With
End Property

'@Description("Gets the number of ticks that represents the date and time of the current DateTimeOffset object in Coordinated Universal Time (UTC).")
Public Property Get UtcTicks() As LongLong
Attribute UtcTicks.VB_Description = "Gets the number of ticks that represents the date and time of the current DateTimeOffset object in Coordinated Universal Time (UTC)."
   UtcTicks = this.DotNetLibDateTimeOffset.UtcTicks()
End Property

'@Description("Gets the year component of the date represented by the current DateTimeOffset object.")
Public Property Get Year() As Long
Attribute Year.VB_Description = "Gets the year component of the date represented by the current DateTimeOffset object."
   Year = this.DotNetLibDateTimeOffset.Year()
End Property

'@Description("Returns a new DateTimeOffset object that adds a specified time interval to the value of this instance.")
Public Function Add(ByVal TimeSpan As TimeSpan) As DateTimeOffset
Attribute Add.VB_Description = "Returns a new DateTimeOffset object that adds a specified time interval to the value of this instance."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.Add(TimeSpan.ComObject)
      Set Add = .Self
   End With
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional days to the value of this instance.")
Public Function AddDays(ByVal Days As Double) As DateTimeOffset
Attribute AddDays.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of whole and fractional days to the value of this instance."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.AddDays(Days)
      Set AddDays = .Self
   End With
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of milliseconds to the value of this instance.")
Public Function AddMilliseconds(ByVal Milliseconds As Double) As DateTimeOffset
Attribute AddMilliseconds.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of milliseconds to the value of this instance."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.AddMilliseconds(Milliseconds)
      Set AddMilliseconds = .Self
   End With
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional minutes to the value of this instance.")
Public Function AddMinutes(ByVal Minutes As Double) As DateTimeOffset
Attribute AddMinutes.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of whole and fractional minutes to the value of this instance."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.AddMinutes(Minutes)
      Set AddMinutes = .Self
   End With
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of months to the value of this instance.")
Public Function AddMonths(ByVal months As Long) As DateTimeOffset
Attribute AddMonths.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of months to the value of this instance."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.AddMonths(months)
      Set AddMonths = .Self
   End With
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional seconds to the value of this instance.")
Public Function AddSeconds(ByVal Seconds As Double) As DateTimeOffset
Attribute AddSeconds.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of whole and fractional seconds to the value of this instance."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.AddSeconds(Seconds)
      Set AddSeconds = .Self
   End With
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of ticks to the value of this instance.")
Public Function AddTicks(ByVal Ticks As LongLong) As DateTimeOffset
Attribute AddTicks.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of ticks to the value of this instance."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.AddTicks(Ticks)
      Set AddTicks = .Self
   End With
End Function

'@Description("Returns a new DateTimeOffset object that adds a specified number of years to the value of this instance.")
Public Function AddYears(ByVal years As Long) As DateTimeOffset
Attribute AddYears.VB_Description = "Returns a new DateTimeOffset object that adds a specified number of years to the value of this instance."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.AddYears(years)
      Set AddYears = .Self
   End With
End Function

'@Description("Compares two DateTimeOffset objects and indicates whether the first is earlier than the second, equal to the second, or later than the second.")
Public Function Compare(ByVal first As DateTimeOffset, ByVal Second As DateTimeOffset) As Long
Attribute Compare.VB_Description = "Compares two DateTimeOffset objects and indicates whether the first is earlier than the second, equal to the second, or later than the second."
   Compare = this.DotNetLibDateTimeOffset.Compare(first.ComObject, Second.ComObject)
End Function

'@Description("Compares the current DateTimeOffset object to a specified DateTimeOffset object and indicates whether the current object is earlier than, the same as, or later than the second DateTimeOffset object.")
Public Function CompareTo(ByVal other As DateTimeOffset) As Long
Attribute CompareTo.VB_Description = "Compares the current DateTimeOffset object to a specified DateTimeOffset object and indicates whether the current object is earlier than, the same as, or later than the second DateTimeOffset object."
   CompareTo = this.DotNetLibDateTimeOffset.CompareTo(other.ComObject)
End Function

'@Description("Determines whether the current DateTimeOffset object represents the same point in time as a specified DateTimeOffset object.")
Public Function Equals(ByVal other As DateTimeOffset) As Boolean
Attribute Equals.VB_Description = "Determines whether the current DateTimeOffset object represents the same point in time as a specified DateTimeOffset object."
   Equals = this.DotNetLibDateTimeOffset.Equals(other.ComObject)
End Function

'@Description("Determines whether a DateTimeOffset object represents the same point in time as a specified object.")
'@Ignore UseMeaningfulName
Public Function Equals2(ByVal obj As Variant) As Boolean
Attribute Equals2.VB_Description = "Determines whether a DateTimeOffset object represents the same point in time as a specified object."
   Equals2 = this.DotNetLibDateTimeOffset.Equals2(obj)
End Function

'@Description("Determines whether two specified DateTimeOffset objects represent the same point in time.")
'@Ignore UseMeaningfulName
Public Function Equals3(ByVal first As DateTimeOffset, ByVal Second As DateTimeOffset) As Boolean
Attribute Equals3.VB_Description = "Determines whether two specified DateTimeOffset objects represent the same point in time."
   Equals3 = this.DotNetLibDateTimeOffset.Equals3(first.ComObject, Second.ComObject)
End Function

'@Description("Determines whether the current DateTimeOffset object represents the same time and has the same offset as a specified DateTimeOffset object.")
Public Function EqualsExact(ByVal other As DateTimeOffset) As Boolean
Attribute EqualsExact.VB_Description = "Determines whether the current DateTimeOffset object represents the same time and has the same offset as a specified DateTimeOffset object."
   EqualsExact = this.DotNetLibDateTimeOffset.EqualsExact(other.ComObject)
End Function

'@Description("Converts the specified Windows file time to an equivalent local time.")
Public Function FromFileTime(ByVal fileTime As LongLong) As DateTimeOffset
Attribute FromFileTime.VB_Description = "Converts the specified Windows file time to an equivalent local time."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.FromFileTime(fileTime)
      Set FromFileTime = .Self
   End With
End Function

'@Description("Converts a Unix time expressed as the number of milliseconds that have elapsed since 1970-01-01T00:00:00Z to a DateTimeOffset value.")
Public Function FromUnixTimeMilliseconds(ByVal Milliseconds As LongLong) As DateTimeOffset
Attribute FromUnixTimeMilliseconds.VB_Description = "Converts a Unix time expressed as the number of milliseconds that have elapsed since 1970-01-01T00:00:00Z to a DateTimeOffset value."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.FromUnixTimeMilliseconds(Milliseconds)
      Set FromUnixTimeMilliseconds = .Self
   End With
End Function

'@Description("Converts a Unix time expressed as the number of seconds that have elapsed since 1970-01-01T00:00:00Z to a DateTimeOffset value.")
Public Function FromUnixTimeSeconds(ByVal Seconds As LongLong) As DateTimeOffset
Attribute FromUnixTimeSeconds.VB_Description = "Converts a Unix time expressed as the number of seconds that have elapsed since 1970-01-01T00:00:00Z to a DateTimeOffset value."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.FromUnixTimeSeconds(Seconds)
      Set FromUnixTimeSeconds = .Self
   End With
End Function

'@Description("Returns the hash code for the current DateTimeOffset object.")
Public Function GetHashCode() As Long
Attribute GetHashCode.VB_Description = "Returns the hash code for the current DateTimeOffset object."
   GetHashCode = this.DotNetLibDateTimeOffset.GetHashCode()
End Function

'@Description("Converts the specified string representation of a date, time, and offset to its DateTimeOffset equivalent.")
Public Function Parse(ByVal strInput As String) As DateTimeOffset
Attribute Parse.VB_Description = "Converts the specified string representation of a date, time, and offset to its DateTimeOffset equivalent."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.Parse(strInput)
      Set Parse = .Self
   End With
End Function

'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified culture-specific format information.")
'@Ignore UseMeaningfulName
Public Function Parse2(ByVal strInput As String, ByVal formatProvider As IFormatProvider) As DateTimeOffset
Attribute Parse2.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified culture-specific format information."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.Parse2(strInput, formatProvider)
      Set Parse2 = .Self
   End With
End Function

'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified culture-specific format information and formatting style.")
'@Ignore UseMeaningfulName
Public Function Parse3(ByVal strInput As String, ByVal formatProvider As IFormatProvider, ByVal styles As DateTimeStyles) As DateTimeOffset
Attribute Parse3.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified culture-specific format information and formatting style."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.Parse3(strInput, formatProvider, styles)
      Set Parse3 = .Self
   End With
End Function

'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")
Public Function ParseExact(ByVal strInput As String, ByVal format As String, ByVal formatProvider As IFormatProvider) As DateTimeOffset
Attribute ParseExact.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.ParseExact(strInput, format, formatProvider)
      Set ParseExact = .Self
   End With
End Function

'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly.")
'@Ignore UseMeaningfulName
Public Function ParseExact2(ByVal strInput As String, ByVal format As String, ByVal formatProvider As IFormatProvider, ByVal styles As DateTimeStyles) As DateTimeOffset
Attribute ParseExact2.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.ParseExact2(strInput, format, formatProvider, styles)
      Set ParseExact2 = .Self
   End With
End Function

'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified formats, culture-specific format information, and style. The format of the string representation must match one of the specified formats exactly.")
'@Ignore UseMeaningfulName
Public Function ParseExact3(ByVal strInput As String, ByRef formats() As String, ByVal formatProvider As IFormatProvider, ByVal styles As DateTimeStyles) As DateTimeOffset
Attribute ParseExact3.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified formats, culture-specific format information, and style. The format of the string representation must match one of the specified formats exactly."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.ParseExact3(strInput, formats, formatProvider, styles)
      Set ParseExact3 = .Self
   End With
End Function

'@Description("Subtracts a DateTimeOffset value that represents a specific date and time from the current DateTimeOffset object.")
Public Function Subtract(ByVal value As DateTimeOffset) As TimeSpan
Attribute Subtract.VB_Description = "Subtracts a DateTimeOffset value that represents a specific date and time from the current DateTimeOffset object."
   With New TimeSpan
      Set .ComObject = this.DotNetLibDateTimeOffset.Subtract(value.ComObject)
      Set Subtract = .Self
   End With
End Function

'@Description("Subtracts a specified time interval from the current DateTimeOffset object.")
'@Ignore UseMeaningfulName
Public Function Subtract2(ByVal value As TimeSpan) As DateTimeOffset
Attribute Subtract2.VB_Description = "Subtracts a specified time interval from the current DateTimeOffset object."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.Subtract2(value.ComObject)
      Set Subtract2 = .Self
   End With
End Function

'@Description("Converts the value of the current DateTimeOffset object to a Windows file time.")
Public Function ToFileTime() As LongLong
Attribute ToFileTime.VB_Description = "Converts the value of the current DateTimeOffset object to a Windows file time."
   ToFileTime = this.DotNetLibDateTimeOffset.ToFileTime()
End Function

'@Description("Converts the current DateTimeOffset object to a DateTimeOffset object that represents the local time.")
Public Function ToLocalTime() As DateTimeOffset
Attribute ToLocalTime.VB_Description = "Converts the current DateTimeOffset object to a DateTimeOffset object that represents the local time."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.ToLocalTime()
      Set ToLocalTime = .Self
   End With
End Function

'@Description("Converts the value of the current DateTimeOffset object to the date and time specified by an offset value.")
Public Function ToOffset(ByVal Offset As TimeSpan) As DateTimeOffset
Attribute ToOffset.VB_Description = "Converts the value of the current DateTimeOffset object to the date and time specified by an offset value."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.ToOffset(Offset.ComObject)
      Set ToOffset = .Self
   End With
End Function

'@DefaultMember
'@Description("Converts the value of the current DateTimeOffset object to its equivalent string representation.")
Public Property Get ToString() As String
Attribute ToString.VB_Description = "Converts the value of the current DateTimeOffset object to its equivalent string representation."
Attribute ToString.VB_UserMemId = 0
   ToString = this.DotNetLibDateTimeOffset.ToString()
End Property

'@Description("Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified format.")
'@Ignore UseMeaningfulName
Public Function ToString2(ByVal format As String) As String
Attribute ToString2.VB_Description = "Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified format."
   ToString2 = this.DotNetLibDateTimeOffset.ToString2(format)
End Function

'@Description("Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified culture-specific formatting information.")
'@Ignore UseMeaningfulName
Public Function ToString3(ByVal formatProvider As IFormatProvider) As String
Attribute ToString3.VB_Description = "Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified culture-specific formatting information."
   ToString3 = this.DotNetLibDateTimeOffset.ToString3(formatProvider)
End Function

'@Description("Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified format and culture-specific format information.")
'@Ignore UseMeaningfulName
Public Function ToString4(ByVal format As String, ByVal formatProvider As IFormatProvider) As String
Attribute ToString4.VB_Description = "Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified format and culture-specific format information."
   ToString4 = this.DotNetLibDateTimeOffset.ToString4(format, formatProvider)
End Function

'@Description("Converts the current DateTimeOffset object to a DateTimeOffset value that represents the Coordinated Universal Time (UTC).")
Public Function ToUniversalTime() As DateTimeOffset
Attribute ToUniversalTime.VB_Description = "Converts the current DateTimeOffset object to a DateTimeOffset value that represents the Coordinated Universal Time (UTC)."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.ToUniversalTime()
      Set ToUniversalTime = .Self
   End With
End Function

'@Description("Returns the number of milliseconds that have elapsed since 1970-01-01T00:00:00.000Z.")
Public Function ToUnixTimeMilliseconds() As LongLong
Attribute ToUnixTimeMilliseconds.VB_Description = "Returns the number of milliseconds that have elapsed since 1970-01-01T00:00:00.000Z."
   ToUnixTimeMilliseconds = this.DotNetLibDateTimeOffset.ToUnixTimeMilliseconds()
End Function

'@Description("Returns the number of seconds that have elapsed since 1970-01-01T00:00:00Z.")
Public Function ToUnixTimeSeconds() As LongLong
Attribute ToUnixTimeSeconds.VB_Description = "Returns the number of seconds that have elapsed since 1970-01-01T00:00:00Z."
   ToUnixTimeSeconds = this.DotNetLibDateTimeOffset.ToUnixTimeSeconds()
End Function

'@Description("Tries to converts a specified string representation of a date and time to its DateTimeOffset equivalent, and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse(ByVal strInput As String, ByRef result As DateTimeOffset) As Boolean
Attribute TryParse.VB_Description = "Tries to converts a specified string representation of a date and time to its DateTimeOffset equivalent, and returns a value that indicates whether the conversion succeeded."
   TryParse = this.DotNetLibDateTimeOffset.TryParse(strInput, result.ComObject)
End Function

'@Description("Tries to convert a specified string representation of a date and time to its DateTimeOffset equivalent, and returns a value that indicates whether the conversion succeeded.")
'@Ignore UseMeaningfulName
Public Function TryParse2(ByVal strInput As String, ByVal formatProvider As IFormatProvider, ByVal styles As DateTimeStyles, ByRef result As DateTimeOffset) As Boolean
Attribute TryParse2.VB_Description = "Tries to convert a specified string representation of a date and time to its DateTimeOffset equivalent, and returns a value that indicates whether the conversion succeeded."
   TryParse2 = this.DotNetLibDateTimeOffset.TryParse2(strInput, formatProvider, styles, result.ComObject)
End Function

'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly.")
Public Function TryParseExact(ByVal strInput As String, ByVal format As String, ByVal formatProvider As IFormatProvider, ByVal styles As DateTimeStyles, ByRef result As DateTimeOffset) As Boolean
Attribute TryParseExact.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly."
   TryParseExact = this.DotNetLibDateTimeOffset.TryParseExact(strInput, format, formatProvider, styles, result.ComObject)
End Function

'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match one of the specified formats exactly.")
'@Ignore UseMeaningfulName
Public Function TryParseExact2(ByVal strInput As String, ByRef formats() As String, ByVal formatProvider As IFormatProvider, ByVal styles As DateTimeStyles, ByRef result As DateTimeOffset) As Boolean
Attribute TryParseExact2.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match one of the specified formats exactly."
   TryParseExact2 = this.DotNetLibDateTimeOffset.TryParseExact2(strInput, formats, formatProvider, styles, result.ComObject)
End Function

'@Description("Adds a specified time interval to a DateTimeOffset object that has a specified date and time, and yields a DateTimeOffset object that has new a date and time.")
Public Function Addition(ByVal DateTimeOffset As DateTimeOffset, ByVal TimeSpan As TimeSpan) As DateTimeOffset
Attribute Addition.VB_Description = "Adds a specified time interval to a DateTimeOffset object that has a specified date and time, and yields a DateTimeOffset object that has new a date and time."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.Addition(DateTimeOffset.ComObject, TimeSpan.ComObject)
      Set Addition = .Self
   End With
End Function

'@Description("Determines whether two specified DateTimeOffset objects represent the same point in time.")
Public Function Equality(ByVal left As DateTimeOffset, ByVal right As DateTimeOffset) As Boolean
Attribute Equality.VB_Description = "Determines whether two specified DateTimeOffset objects represent the same point in time."
   Equality = this.DotNetLibDateTimeOffset.Equality(left.ComObject, right.ComObject)
End Function

'@Description("Determines whether one specified DateTimeOffset object is greater than (or later than) a second specified DateTimeOffset object.")
Public Function GreaterThan(ByVal left As DateTimeOffset, ByVal right As DateTimeOffset) As Boolean
Attribute GreaterThan.VB_Description = "Determines whether one specified DateTimeOffset object is greater than (or later than) a second specified DateTimeOffset object."
   GreaterThan = this.DotNetLibDateTimeOffset.GreaterThan(left.ComObject, right.ComObject)
End Function

'@Description("Determines whether one specified DateTimeOffset object is greater than or equal to a second specified DateTimeOffset object.")
Public Function GreaterThanOrEqual(ByVal left As DateTimeOffset, ByVal right As DateTimeOffset) As Boolean
Attribute GreaterThanOrEqual.VB_Description = "Determines whether one specified DateTimeOffset object is greater than or equal to a second specified DateTimeOffset object."
   GreaterThanOrEqual = this.DotNetLibDateTimeOffset.GreaterThanOrEqual(left.ComObject, right.ComObject)
End Function

'@Description("Defines an implicit conversion of a DateTime object to a DateTimeOffset object.")
Public Function Implicit(ByVal DateTime As DateTime) As DateTimeOffset
Attribute Implicit.VB_Description = "Defines an implicit conversion of a DateTime object to a DateTimeOffset object."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.Implicit(DateTime.ComObject)
      Set Implicit = .Self
   End With
End Function

'@Description("Determines whether two specified DateTimeOffset objects refer to different points in time.")
Public Function Inequality(ByVal left As DateTimeOffset, ByVal right As DateTimeOffset) As Boolean
Attribute Inequality.VB_Description = "Determines whether two specified DateTimeOffset objects refer to different points in time."
   Inequality = this.DotNetLibDateTimeOffset.Inequality(left.ComObject, right.ComObject)
End Function

'@Description("Determines whether one specified DateTimeOffset object is less than a second specified DateTimeOffset object.")
Public Function LessThan(ByVal left As DateTimeOffset, ByVal right As DateTimeOffset) As Boolean
Attribute LessThan.VB_Description = "Determines whether one specified DateTimeOffset object is less than a second specified DateTimeOffset object."
   LessThan = this.DotNetLibDateTimeOffset.LessThan(left.ComObject, right.ComObject)
End Function

'@Description("Determines whether one specified DateTimeOffset object is less than or equal to a second specified DateTimeOffset object.")
Public Function LessThanOrEqual(ByVal left As DateTimeOffset, ByVal right As DateTimeOffset) As Boolean
Attribute LessThanOrEqual.VB_Description = "Determines whether one specified DateTimeOffset object is less than or equal to a second specified DateTimeOffset object."
   LessThanOrEqual = this.DotNetLibDateTimeOffset.LessThanOrEqual(left.ComObject, right.ComObject)
End Function

'@Description("Subtracts one DateTimeOffset object from another and yields a time interval.")
Public Function Subtraction(ByVal left As DateTimeOffset, ByVal right As DateTimeOffset) As TimeSpan
Attribute Subtraction.VB_Description = "Subtracts one DateTimeOffset object from another and yields a time interval."
   With New TimeSpan
      Set .ComObject = this.DotNetLibDateTimeOffset.Subtraction(left.ComObject, right.ComObject)
      Set Subtraction = .Self
   End With
End Function

'@Description("Subtracts a specified time interval from a specified date and time, and yields a new date and time.")
'@Ignore UseMeaningfulName
Public Function Subtraction2(ByVal DateTimeOffset As DateTimeOffset, ByVal TimeSpan As TimeSpan) As DateTimeOffset
Attribute Subtraction2.VB_Description = "Subtracts a specified time interval from a specified date and time, and yields a new date and time."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibDateTimeOffset.Subtraction2(DateTimeOffset.ComObject, TimeSpan.ComObject)
      Set Subtraction2 = .Self
   End With
End Function
