VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "DateTimeOffset factory methods and static members."
'@Folder "VBADotNetLib.System"
'@ModuleDescription("DateTimeOffset factory methods and static members.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v2.0 August 4, 2023
'@LastModified September 16, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1

'@Remarks
' Singleton class for DateTimeOffset containing constructors and static members.

Option Explicit

Private Type TDateTimeOffset
    DateTimeOffsetSingleton As DotNetLib.DateTimeOffsetSingleton
End Type

Private this As TDateTimeOffset

Private Sub Class_Initialize()
    Set this.DateTimeOffsetSingleton = New DotNetLib.DateTimeOffsetSingleton
End Sub

Private Sub Class_Terminate()
    Set this.DateTimeOffsetSingleton = Nothing
End Sub

'--------------------------------------------------------------------
'Constructors
'--------------------------------------------------------------------

'@Static
'@Description("Initializes a new instance of the DateTimeOffset structure using the specified DateTime value.")
Public Function CreateFromDateTime(ByVal pDateTime As DotNetLib.DateTime) As DotNetLib.DateTimeOffset
Attribute CreateFromDateTime.VB_Description = "Initializes a new instance of the DateTimeOffset structure using the specified DateTime value."
    Set CreateFromDateTime = this.DateTimeOffsetSingleton.CreateFromDateTime(pDateTime)
End Function

'@Static
'@Description("Initializes a new instance of the DateTimeOffset structure using the specified DateTime value and offset.")
Public Function CreateFromDateTime2(ByVal pDateTime As DotNetLib.DateTime, ByVal pOffset As DotNetLib.TimeSpan) As DotNetLib.DateTimeOffset
Attribute CreateFromDateTime2.VB_Description = "Initializes a new instance of the DateTimeOffset structure using the specified DateTime value and offset."
    Set CreateFromDateTime2 = this.DateTimeOffsetSingleton.CreateFromDateTime2(pDateTime, pOffset)
End Function

'@Static
'@Description("Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, and offset.")
Public Function CreateFromDateTimeParts(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long, ByVal pOffset As DotNetLib.TimeSpan) As DotNetLib.DateTimeOffset
Attribute CreateFromDateTimeParts.VB_Description = "Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, and offset."
    Set CreateFromDateTimeParts = this.DateTimeOffsetSingleton.CreateFromDateTimeParts(pYear, pMonth, pDay, pHour, pMinute, pSecond, pOffset)
End Function

'@Description("Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, millisecond, and offset.")
Public Function CreateFromDateTimeParts2(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long, ByVal pMillisecond As Long, ByVal pOffset As DotNetLib.TimeSpan) As DotNetLib.DateTimeOffset
Attribute CreateFromDateTimeParts2.VB_Description = "Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, millisecond, and offset."
    Set CreateFromDateTimeParts2 = this.DateTimeOffsetSingleton.CreateFromDateTimeParts2(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond, pOffset)
End Function

'@Static
'@Description("Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, millisecond, and offset of a specified calendar.")
Public Function CreateFromDateTimeParts3(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long, ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long, ByVal pMillisecond As Long, ByVal pCalendar As DotNetLib.Calendar, ByVal pOffset As DotNetLib.TimeSpan) As DotNetLib.DateTimeOffset
Attribute CreateFromDateTimeParts3.VB_Description = "Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, millisecond, and offset of a specified calendar."
    Set CreateFromDateTimeParts3 = this.DateTimeOffsetSingleton.CreateFromDateTimeParts3(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond, pCalendar, pOffset)
End Function

'@Static
'@Description("Initializes a new instance of the DateTimeOffset structure using the specified number of ticks and offset.")
Public Function CreateFromTicks(ByVal pTicks As LongLong, ByVal pOffset As DotNetLib.TimeSpan) As DotNetLib.DateTimeOffset
Attribute CreateFromTicks.VB_Description = "Initializes a new instance of the DateTimeOffset structure using the specified number of ticks and offset."
    Set CreateFromTicks = this.DateTimeOffsetSingleton.CreateFromTicks(pTicks, pOffset)
End Function

'--------------------------------------------------------------------
'Fields
'--------------------------------------------------------------------

'@Static
'@Description("Represents the greatest possible value of DateTimeOffset. This field is read-only.")
Public Property Get MaxValue() As DotNetLib.DateTimeOffset
Attribute MaxValue.VB_Description = "Represents the greatest possible value of DateTimeOffset. This field is read-only."
   Set MaxValue = this.DateTimeOffsetSingleton.MaxValue
End Property

'@Static
'@Description("Represents the earliest possible DateTimeOffset value. This field is read-only.")
Public Property Get MinValue() As DotNetLib.DateTimeOffset
Attribute MinValue.VB_Description = "Represents the earliest possible DateTimeOffset value. This field is read-only."
   Set MinValue = this.DateTimeOffsetSingleton.MinValue
End Property

'--------------------------------------------------------------------
'Properties
'--------------------------------------------------------------------

'@Static
'@Description("Gets a DateTimeOffset object that is set to the current date and time on the current computer, with the offset set to the local time's offset from Coordinated Universal Time (UTC).")
Public Property Get Now() As DotNetLib.DateTimeOffset
Attribute Now.VB_Description = "Gets a DateTimeOffset object that is set to the current date and time on the current computer, with the offset set to the local time's offset from Coordinated Universal Time (UTC)."
    Set Now = this.DateTimeOffsetSingleton.Now()
End Property

'@Static
'@Description("Gets a DateTimeOffset object whose date and time are set to the current Coordinated Universal Time (UTC) date and time and whose offset is Zero.")
Public Property Get UtcNow() As DotNetLib.DateTimeOffset
Attribute UtcNow.VB_Description = "Gets a DateTimeOffset object whose date and time are set to the current Coordinated Universal Time (UTC) date and time and whose offset is Zero."
    Set UtcNow = this.DateTimeOffsetSingleton.UtcNow()
End Property

'--------------------------------------------------------------------
'Methods
'--------------------------------------------------------------------

'@Static
'@Description("Compares two DateTimeOffset objects and indicates whether the first is earlier than the second, equal to the second, or later than the second.")
Public Function Compare(ByVal FIRST As DotNetLib.DateTimeOffset, ByVal SECOND As DotNetLib.DateTimeOffset) As Long
Attribute Compare.VB_Description = "Compares two DateTimeOffset objects and indicates whether the first is earlier than the second, equal to the second, or later than the second."
    Compare = this.DateTimeOffsetSingleton.Compare(FIRST, SECOND)
End Function

'@Static
'@Description("Determines whether two specified DateTimeOffset objects represent the same point in time.")
Public Function Equals(ByVal FIRST As DotNetLib.DateTimeOffset, ByVal SECOND As DotNetLib.DateTimeOffset) As Boolean
Attribute Equals.VB_Description = "Determines whether two specified DateTimeOffset objects represent the same point in time."
    Equals = this.DateTimeOffsetSingleton.Equals(FIRST, SECOND)
End Function

'@Description("Determines whether the current DateTimeOffset object represents the same time and has the same offset as a specified DateTimeOffset object.")
Public Function EqualsExact(ByVal other As DotNetLib.DateTimeOffset) As Boolean
Attribute EqualsExact.VB_Description = "Determines whether the current DateTimeOffset object represents the same time and has the same offset as a specified DateTimeOffset object."
    EqualsExact = this.DateTimeOffsetSingleton.EqualsExact(other)
End Function

'@Static
'@Description("Converts the specified Windows file time to an equivalent local time.")
Public Function FromFileTime(ByVal fileTime As LongLong) As DotNetLib.DateTimeOffset
Attribute FromFileTime.VB_Description = "Converts the specified Windows file time to an equivalent local time."
    Set FromFileTime = this.DateTimeOffsetSingleton.FromFileTime(fileTime)
End Function

'@Static
'@Description("Converts a Unix time expressed as the number of milliseconds that have elapsed since 1970-01-01T00:00:00Z to a DateTimeOffset value.")
Public Function FromUnixTimeMilliseconds(ByVal pMilliseconds As LongLong) As DotNetLib.DateTimeOffset
Attribute FromUnixTimeMilliseconds.VB_Description = "Converts a Unix time expressed as the number of milliseconds that have elapsed since 1970-01-01T00:00:00Z to a DateTimeOffset value."
    Set FromUnixTimeMilliseconds = this.DateTimeOffsetSingleton.FromUnixTimeMilliseconds(pMilliseconds)
End Function

'@Static
'@Description("Converts a Unix time expressed as the number of seconds that have elapsed since 1970-01-01T00:00:00Z to a DateTimeOffset value.")
Public Function FromUnixTimeSeconds(ByVal Seconds As LongLong) As DotNetLib.DateTimeOffset
Attribute FromUnixTimeSeconds.VB_Description = "Converts a Unix time expressed as the number of seconds that have elapsed since 1970-01-01T00:00:00Z to a DateTimeOffset value."
    Set FromUnixTimeSeconds = this.DateTimeOffsetSingleton.FromUnixTimeSeconds(Seconds)
End Function

'@Static
'@Description("Converts the specified string representation of a date, time, and offset to its DateTimeOffset equivalent.")
Public Function Parse(ByVal strInput As String) As DotNetLib.DateTimeOffset
Attribute Parse.VB_Description = "Converts the specified string representation of a date, time, and offset to its DateTimeOffset equivalent."
    Set Parse = this.DateTimeOffsetSingleton.Parse(strInput)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified culture-specific format information.")
Public Function Parse2(ByVal strInput As String, ByVal formatProvider As IFormatProvider) As DotNetLib.DateTimeOffset
Attribute Parse2.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified culture-specific format information."
    Set Parse2 = this.DateTimeOffsetSingleton.Parse2(strInput, formatProvider)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified culture-specific format information and formatting style.")
Public Function Parse3(ByVal strInput As String, ByVal formatProvider As IFormatProvider, ByVal styles As DateTimeStyles) As DotNetLib.DateTimeOffset
Attribute Parse3.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified culture-specific format information and formatting style."
    Set Parse3 = this.DateTimeOffsetSingleton.Parse3(strInput, formatProvider, styles)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")
Public Function ParseExact(ByVal strInput As String, ByVal Format As String, ByVal formatProvider As IFormatProvider) As DotNetLib.DateTimeOffset
Attribute ParseExact.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly."
    Set ParseExact = this.DateTimeOffsetSingleton.ParseExact(strInput, Format, formatProvider)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly.")
Public Function ParseExact2(ByVal strInput As String, ByVal Format As String, ByVal formatProvider As IFormatProvider, ByVal styles As DateTimeStyles) As DotNetLib.DateTimeOffset
Attribute ParseExact2.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly."
    Set ParseExact2 = this.DateTimeOffsetSingleton.ParseExact2(strInput, Format, formatProvider, styles)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified formats, culture-specific format information, and style. The format of the string representation must match one of the specified formats exactly.")
Public Function ParseExact3(ByVal strInput As String, ByRef formats() As String, ByVal formatProvider As IFormatProvider, ByVal styles As DateTimeStyles) As DotNetLib.DateTimeOffset
Attribute ParseExact3.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified formats, culture-specific format information, and style. The format of the string representation must match one of the specified formats exactly."
    Set ParseExact3 = this.DateTimeOffsetSingleton.ParseExact3(strInput, formats, formatProvider, styles)
End Function

'@Static
'@Description("Tries to converts a specified string representation of a date and time to its DateTimeOffset equivalent, and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse(ByVal strInput As String, ByRef result As DotNetLib.DateTimeOffset) As Boolean
Attribute TryParse.VB_Description = "Tries to converts a specified string representation of a date and time to its DateTimeOffset equivalent, and returns a value that indicates whether the conversion succeeded."
    TryParse = this.DateTimeOffsetSingleton.TryParse(strInput, result)
End Function

'@Static
'@Description("Tries to convert a specified string representation of a date and time to its DateTimeOffset equivalent, and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse2(ByVal strInput As String, ByVal formatProvider As IFormatProvider, ByVal styles As DateTimeStyles, ByRef result As DotNetLib.DateTimeOffset) As Boolean
Attribute TryParse2.VB_Description = "Tries to convert a specified string representation of a date and time to its DateTimeOffset equivalent, and returns a value that indicates whether the conversion succeeded."
    TryParse2 = this.DateTimeOffsetSingleton.TryParse2(strInput, formatProvider, styles, result)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly.")
Public Function TryParseExact(ByVal strInput As String, ByVal Format As String, ByVal formatProvider As IFormatProvider, ByVal styles As DateTimeStyles, ByRef result As DotNetLib.DateTimeOffset) As Boolean
Attribute TryParseExact.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly."
    TryParseExact = this.DateTimeOffsetSingleton.TryParseExact(strInput, Format, formatProvider, styles, result)
End Function

'@Static
'@Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match one of the specified formats exactly.")
Public Function TryParseExact2(ByVal strInput As String, ByRef formats() As String, ByVal formatProvider As IFormatProvider, ByVal styles As DateTimeStyles, ByRef result As DotNetLib.DateTimeOffset) As Boolean
Attribute TryParseExact2.VB_Description = "Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match one of the specified formats exactly."
    TryParseExact2 = this.DateTimeOffsetSingleton.TryParseExact2(strInput, formats, formatProvider, styles, result)
End Function

'--------------------------------------------------------------------
'Operators
'--------------------------------------------------------------------

'@Static
'@Description("Adds a specified time interval to a DateTimeOffset object that has a specified date and time, and yields a DateTimeOffset object that has new a date and time.")
Public Function Addition(ByVal pDateTimeOffset As DotNetLib.DateTimeOffset, ByVal pTimeSpan As DotNetLib.TimeSpan) As DotNetLib.DateTimeOffset
Attribute Addition.VB_Description = "Adds a specified time interval to a DateTimeOffset object that has a specified date and time, and yields a DateTimeOffset object that has new a date and time."
    Set Addition = this.DateTimeOffsetSingleton.Addition(pDateTimeOffset, pTimeSpan)
End Function

'@Static
'@Description("Determines whether two specified DateTimeOffset objects represent the same point in time.")
Public Function Equality(ByVal Left As DotNetLib.DateTimeOffset, ByVal Right As DotNetLib.DateTimeOffset) As Boolean
Attribute Equality.VB_Description = "Determines whether two specified DateTimeOffset objects represent the same point in time."
   Equality = this.DateTimeOffsetSingleton.Equality(Left, Right)
End Function

'@Static
'@Description("Determines whether one specified DateTimeOffset object is greater than (or later than) a second specified DateTimeOffset object.")
Public Function GreaterThan(ByVal Left As DotNetLib.DateTimeOffset, ByVal Right As DotNetLib.DateTimeOffset) As Boolean
Attribute GreaterThan.VB_Description = "Determines whether one specified DateTimeOffset object is greater than (or later than) a second specified DateTimeOffset object."
    GreaterThan = this.DateTimeOffsetSingleton.GreaterThan(Left, Right)
End Function

'@Static
'@Description("Determines whether one specified DateTimeOffset object is greater than or equal to a second specified DateTimeOffset object.")
Public Function GreaterThanOrEqual(ByVal Left As DotNetLib.DateTimeOffset, ByVal Right As DotNetLib.DateTimeOffset) As Boolean
Attribute GreaterThanOrEqual.VB_Description = "Determines whether one specified DateTimeOffset object is greater than or equal to a second specified DateTimeOffset object."
    GreaterThanOrEqual = this.DateTimeOffsetSingleton.GreaterThanOrEqual(Left, Right)
End Function

'@Static
'@Description("Defines an implicit conversion of a DateTime object to a DateTimeOffset object.")
Public Function Implicit(ByVal DateTime As DotNetLib.DateTime) As DotNetLib.DateTimeOffset
Attribute Implicit.VB_Description = "Defines an implicit conversion of a DateTime object to a DateTimeOffset object."
    Set Implicit = this.DateTimeOffsetSingleton.Implicit(DateTime)
End Function

'@Static
'@Description("Determines whether two specified DateTimeOffset objects refer to different points in time.")
Public Function Inequality(ByVal Left As DotNetLib.DateTimeOffset, ByVal Right As DotNetLib.DateTimeOffset) As Boolean
Attribute Inequality.VB_Description = "Determines whether two specified DateTimeOffset objects refer to different points in time."
    Inequality = this.DateTimeOffsetSingleton.Inequality(Left, Right)
End Function

'@Static
'@Description("Determines whether one specified DateTimeOffset object is less than a second specified DateTimeOffset object.")
Public Function LessThan(ByVal Left As DotNetLib.DateTimeOffset, ByVal Right As DotNetLib.DateTimeOffset) As Boolean
Attribute LessThan.VB_Description = "Determines whether one specified DateTimeOffset object is less than a second specified DateTimeOffset object."
   LessThan = this.DateTimeOffsetSingleton.LessThan(Left, Right)
End Function

'@Static
'@Description("Determines whether one specified DateTimeOffset object is less than or equal to a second specified DateTimeOffset object.")
Public Function LessThanOrEqual(ByVal Left As DotNetLib.DateTimeOffset, ByVal Right As DotNetLib.DateTimeOffset) As Boolean
Attribute LessThanOrEqual.VB_Description = "Determines whether one specified DateTimeOffset object is less than or equal to a second specified DateTimeOffset object."
    LessThanOrEqual = this.DateTimeOffsetSingleton.LessThanOrEqual(Left, Right)
End Function

'@Static
'@Description("Subtracts one DateTimeOffset object from another and yields a time interval.")
Public Function Subtraction(ByVal Left As DotNetLib.DateTimeOffset, ByVal Right As DotNetLib.DateTimeOffset) As DotNetLib.TimeSpan
Attribute Subtraction.VB_Description = "Subtracts one DateTimeOffset object from another and yields a time interval."
    Set Subtraction = this.DateTimeOffsetSingleton.Subtraction(Left, Right)
End Function

'@Static
'@Description("Subtracts a specified time interval from a specified date and time, and yields a new date and time.")
Public Function Subtraction2(ByVal DateTimeOffset As DotNetLib.DateTimeOffset, ByVal TimeSpan As DotNetLib.TimeSpan) As DotNetLib.DateTimeOffset
Attribute Subtraction2.VB_Description = "Subtracts a specified time interval from a specified date and time, and yields a new date and time."
    Set Subtraction2 = this.DateTimeOffsetSingleton.Subtraction2(DateTimeOffset, TimeSpan)
End Function
