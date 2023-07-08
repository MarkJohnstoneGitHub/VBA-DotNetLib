VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeSpan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a time interval."
'@Folder("DotNet.System")
'@PredeclaredId
'@ModuleDescription("Represents a time interval.")

Option Explicit

Private Type TTimeSpan
   DotNetLibTimeSpan As DotNetLib.TimeSpan
End Type

Private this As TTimeSpan

Private Sub Class_Initialize()
   Set this.DotNetLibTimeSpan = New DotNetLib.TimeSpan
End Sub

Private Sub Class_Terminate()
   Set this.DotNetLibTimeSpan = Nothing
End Sub

Friend Property Get TimeSpanComObject() As DotNetLib.TimeSpan
   Set TimeSpanComObject = this.DotNetLibTimeSpan
End Property

Friend Property Set TimeSpanComObject(ByVal objTimeSpan As DotNetLib.TimeSpan)
   Set this.DotNetLibTimeSpan = objTimeSpan
End Property

Friend Property Get Self() As TimeSpan
   Set Self = Me
End Property

'@Description("Initializes a new instance of the TimeSpan structure to the specified number of ticks.")
Public Function CreateFromTicks(ByVal Ticks As LongLong) As TimeSpan
Attribute CreateFromTicks.VB_Description = "Initializes a new instance of the TimeSpan structure to the specified number of ticks."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.CreateFromTicks(Ticks)
      Set CreateFromTicks = .Self
   End With
End Function

'@Description("Initializes a new instance of the TimeSpan structure to a specified number of hours, minutes, and seconds.")
Public Function Create(ByVal Hours As Long, ByVal Minutes As Long, ByVal Seconds As Long) As TimeSpan
Attribute Create.VB_Description = "Initializes a new instance of the TimeSpan structure to a specified number of hours, minutes, and seconds."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.Create(Hours, Minutes, Seconds)
      Set Create = .Self
   End With
End Function

'@Description("Initializes a new instance of the TimeSpan structure to a specified number of days, hours, minutes, and seconds.")
Public Function Create2(ByVal Days As Long, ByVal Hours As Long, ByVal Minutes As Long, ByVal Seconds As Long) As TimeSpan
Attribute Create2.VB_Description = "Initializes a new instance of the TimeSpan structure to a specified number of days, hours, minutes, and seconds."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.Create2(Days, Hours, Minutes, Seconds)
      Set Create2 = .Self
   End With
End Function

'@Description("Initializes a new instance of the TimeSpan structure to a specified number of days, hours, minutes, seconds, and milliseconds.")
Public Function Create3(ByVal Days As Long, ByVal Hours As Long, ByVal Minutes As Long, ByVal Seconds As Long, ByVal Milliseconds As Long) As TimeSpan
Attribute Create3.VB_Description = "Initializes a new instance of the TimeSpan structure to a specified number of days, hours, minutes, seconds, and milliseconds."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.Create3(Days, Hours, Minutes, Seconds, Milliseconds)
      Set Create3 = .Self
   End With
End Function

'@StaticField
'@Description("Represents the maximum TimeSpan value. This field is read-only.")
Public Property Get MaxValue() As TimeSpan
Attribute MaxValue.VB_Description = "Represents the maximum TimeSpan value. This field is read-only."
   Set MaxValue = TimeSpanStatic.MaxValue
End Property

'@StaticField
'@Description("Represents the minimum TimeSpan value. This field is read-only.")
Public Property Get MinValue() As TimeSpan
Attribute MinValue.VB_Description = "Represents the minimum TimeSpan value. This field is read-only."
   Set MinValue = TimeSpanStatic.MinValue
End Property

'@StaticField
'@Description("Represents the number of ticks in 1 day. This field is constant.")
Public Property Get TicksPerDay() As LongLong
Attribute TicksPerDay.VB_Description = "Represents the number of ticks in 1 day. This field is constant."
   TicksPerDay = TimeSpanStatic.TicksPerDay
End Property

'@StaticField
'@Description("Represents the number of ticks in 1 hour. This field is constant.")
Public Property Get TicksPerHour() As LongLong
Attribute TicksPerHour.VB_Description = "Represents the number of ticks in 1 hour. This field is constant."
   TicksPerHour = TimeSpanStatic.TicksPerHour()
End Property

'@StaticField
'@Description("Represents the number of ticks in 1 millisecond. This field is constant.")
Public Property Get TicksPerMillisecond() As LongLong
Attribute TicksPerMillisecond.VB_Description = "Represents the number of ticks in 1 millisecond. This field is constant."
   TicksPerMillisecond = TimeSpanStatic.TicksPerMillisecond()
End Property

'@StaticField
'@Description("Represents the number of ticks in 1 minute. This field is constant.")
Public Property Get TicksPerMinute() As LongLong
Attribute TicksPerMinute.VB_Description = "Represents the number of ticks in 1 minute. This field is constant."
   TicksPerMinute = TimeSpanStatic.TicksPerMinute()
End Property

'@Description("Represents the number of ticks in 1 second.")
Public Property Get TicksPerSecond() As LongLong
Attribute TicksPerSecond.VB_Description = "Represents the number of ticks in 1 second."
   TicksPerSecond = this.DotNetLibTimeSpan.TicksPerSecond()
End Property

'@Description("Represents the zero TimeSpan value. This field is read-only.")
Public Property Get Zero() As TimeSpan
Attribute Zero.VB_Description = "Represents the zero TimeSpan value. This field is read-only."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.Zero()
      Set Zero = .Self
   End With
End Property

'@Description("Gets the days component of the time interval represented by the current TimeSpan structure.")
Public Property Get Days() As Long
Attribute Days.VB_Description = "Gets the days component of the time interval represented by the current TimeSpan structure."
   Days = this.DotNetLibTimeSpan.Days()
End Property

'@Description("Gets the hours component of the time interval represented by the current TimeSpan structure.")
Public Property Get Hours() As Long
Attribute Hours.VB_Description = "Gets the hours component of the time interval represented by the current TimeSpan structure."
   Hours = this.DotNetLibTimeSpan.Hours()
End Property

'@Description("Gets the milliseconds component of the time interval represented by the current TimeSpan structure.")
Public Property Get Milliseconds() As Long
Attribute Milliseconds.VB_Description = "Gets the milliseconds component of the time interval represented by the current TimeSpan structure."
   Milliseconds = this.DotNetLibTimeSpan.Milliseconds()
End Property

'@Description("Gets the minutes component of the time interval represented by the current TimeSpan structure.")
Public Property Get Minutes() As Long
Attribute Minutes.VB_Description = "Gets the minutes component of the time interval represented by the current TimeSpan structure."
   Minutes = this.DotNetLibTimeSpan.Minutes()
End Property

'@Description("Gets the seconds component of the time interval represented by the current TimeSpan structure.")
Public Property Get Seconds() As Long
Attribute Seconds.VB_Description = "Gets the seconds component of the time interval represented by the current TimeSpan structure."
   Seconds = this.DotNetLibTimeSpan.Seconds()
End Property

'@Description("Gets the number of ticks that represent the value of the current TimeSpan structure.")
Public Property Get Ticks() As LongLong
Attribute Ticks.VB_Description = "Gets the number of ticks that represent the value of the current TimeSpan structure."
   Ticks = this.DotNetLibTimeSpan.Ticks()
End Property

'@Description("Gets the value of the current TimeSpan structure expressed in whole and fractional days.")
Public Property Get TotalDays() As Double
Attribute TotalDays.VB_Description = "Gets the value of the current TimeSpan structure expressed in whole and fractional days."
   TotalDays = this.DotNetLibTimeSpan.TotalDays()
End Property

'@Description("Gets the value of the current TimeSpan structure expressed in whole and fractional hours.")
Public Property Get TotalHours() As Double
Attribute TotalHours.VB_Description = "Gets the value of the current TimeSpan structure expressed in whole and fractional hours."
   TotalHours = this.DotNetLibTimeSpan.TotalHours()
End Property

'@Description("Gets the value of the current TimeSpan structure expressed in whole and fractional milliseconds.")
Public Property Get TotalMilliseconds() As Double
Attribute TotalMilliseconds.VB_Description = "Gets the value of the current TimeSpan structure expressed in whole and fractional milliseconds."
   TotalMilliseconds = this.DotNetLibTimeSpan.TotalMilliseconds()
End Property

'@Description("Gets the value of the current TimeSpan structure expressed in whole and fractional minutes.")
Public Property Get TotalMinutes() As Double
Attribute TotalMinutes.VB_Description = "Gets the value of the current TimeSpan structure expressed in whole and fractional minutes."
   TotalMinutes = this.DotNetLibTimeSpan.TotalMinutes()
End Property

'@Description("Gets the value of the current TimeSpan structure expressed in whole and fractional seconds.")
Public Property Get TotalSeconds() As Double
Attribute TotalSeconds.VB_Description = "Gets the value of the current TimeSpan structure expressed in whole and fractional seconds."
   TotalSeconds = this.DotNetLibTimeSpan.TotalSeconds()
End Property

'@Description("Returns a new TimeSpan object whose value is the sum of the specified TimeSpan object and this instance.")
Public Function Add(ByRef ts As TimeSpan) As TimeSpan
Attribute Add.VB_Description = "Returns a new TimeSpan object whose value is the sum of the specified TimeSpan object and this instance."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.Add(ts.TimeSpanComObject)
      Set Add = .Self
   End With
End Function

'@Description("Compares two TimeSpan values and returns an integer that indicates whether the first value is shorter than, equal to, or longer than the second value.")
Public Function Compare(ByRef t1 As TimeSpan, ByRef t2 As TimeSpan) As Long
Attribute Compare.VB_Description = "Compares two TimeSpan values and returns an integer that indicates whether the first value is shorter than, equal to, or longer than the second value."
   Compare = this.DotNetLibTimeSpan.Compare(t1.TimeSpanComObject, t2.TimeSpanComObject)
End Function

'@Description("Compares this instance to a specified TimeSpan object and returns an integer that indicates whether this instance is shorter than, equal to, or longer than the TimeSpan object.")
Public Function CompareTo(ByRef value As TimeSpan) As Long
Attribute CompareTo.VB_Description = "Compares this instance to a specified TimeSpan object and returns an integer that indicates whether this instance is shorter than, equal to, or longer than the TimeSpan object."
   CompareTo = this.DotNetLibTimeSpan.CompareTo(value.TimeSpanComObject)
End Function

'@Description("Compares this instance to a specified object and returns an integer that indicates whether this instance is shorter than, equal to, or longer than the specified object.")
Public Function CompareTo2(ByVal value As Variant) As Long
Attribute CompareTo2.VB_Description = "Compares this instance to a specified object and returns an integer that indicates whether this instance is shorter than, equal to, or longer than the specified object."
   CompareTo2 = this.DotNetLibTimeSpan.CompareTo2(value)
End Function

'@Description("Returns a new TimeSpan object whose value is the absolute value of the current TimeSpan object.")
Public Function Duration() As TimeSpan
Attribute Duration.VB_Description = "Returns a new TimeSpan object whose value is the absolute value of the current TimeSpan object."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.Duration()
      Set Duration = .Self
   End With
End Function

'@Description("Returns a value indicating whether this instance is equal to a specified TimeSpan object.")
Public Function Equals(ByRef obj As TimeSpan) As Boolean
Attribute Equals.VB_Description = "Returns a value indicating whether this instance is equal to a specified TimeSpan object."
   Equals = this.DotNetLibTimeSpan.Equals(obj.TimeSpanComObject)
End Function

'@Description("Returns a value indicating whether this instance is equal to a specified object.")
Public Function Equals2(ByVal value As Variant) As Boolean
Attribute Equals2.VB_Description = "Returns a value indicating whether this instance is equal to a specified object."
   Equals2 = this.DotNetLibTimeSpan.Equals2(value)
End Function

'@Description("Returns a value that indicates whether two specified instances of TimeSpan are equal.")
Public Function Equals3(ByRef t1 As TimeSpan, ByRef t2 As TimeSpan) As Boolean
Attribute Equals3.VB_Description = "Returns a value that indicates whether two specified instances of TimeSpan are equal."
   Equals3 = this.DotNetLibTimeSpan.Equals3(t1.TimeSpanComObject, t2.TimeSpanComObject)
End Function

'@Description("Returns a TimeSpan that represents a specified number of days, where the specification is accurate to the nearest millisecond.")
Public Function FromDays(ByVal value As Double) As TimeSpan
Attribute FromDays.VB_Description = "Returns a TimeSpan that represents a specified number of days, where the specification is accurate to the nearest millisecond."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.FromDays(value)
      Set FromDays = .Self
   End With
End Function

'@Description("Returns a TimeSpan that represents a specified number of hours, where the specification is accurate to the nearest millisecond.")
Public Function FromHours(ByVal value As Double) As TimeSpan
Attribute FromHours.VB_Description = "Returns a TimeSpan that represents a specified number of hours, where the specification is accurate to the nearest millisecond."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.FromHours(value)
      Set FromHours = .Self
   End With
End Function

'@Description("Returns a TimeSpan that represents a specified number of milliseconds.")
Public Function FromMilliseconds(ByVal value As Double) As TimeSpan
Attribute FromMilliseconds.VB_Description = "Returns a TimeSpan that represents a specified number of milliseconds."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.FromMilliseconds(value)
      Set FromMilliseconds = .Self
   End With
End Function

'@Description("Returns a TimeSpan that represents a specified number of minutes, where the specification is accurate to the nearest millisecond.")
Public Function FromMinutes(ByVal value As Double) As TimeSpan
Attribute FromMinutes.VB_Description = "Returns a TimeSpan that represents a specified number of minutes, where the specification is accurate to the nearest millisecond."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.FromMinutes(value)
      Set FromMinutes = .Self
   End With
End Function

'@Description("Returns a TimeSpan that represents a specified number of seconds, where the specification is accurate to the nearest millisecond.")
Public Function FromSeconds(ByVal value As Double) As TimeSpan
Attribute FromSeconds.VB_Description = "Returns a TimeSpan that represents a specified number of seconds, where the specification is accurate to the nearest millisecond."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.FromSeconds(value)
      Set FromSeconds = .Self
   End With
End Function

'@Description("Returns a TimeSpan that represents a specified time, where the specification is in units of ticks.")
Public Function FromTicks(ByVal value As LongLong) As TimeSpan
Attribute FromTicks.VB_Description = "Returns a TimeSpan that represents a specified time, where the specification is in units of ticks."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.FromTicks(value)
      Set FromTicks = .Self
   End With
End Function

'@Description("Returns a hash code for this instance.")
Public Function GetHashCode() As Long
Attribute GetHashCode.VB_Description = "Returns a hash code for this instance."
   GetHashCode = this.DotNetLibTimeSpan.GetHashCode()
End Function

'@Description("Returns a new TimeSpan object whose value is the negated value of this instance.")
Public Function Negate() As TimeSpan
Attribute Negate.VB_Description = "Returns a new TimeSpan object whose value is the negated value of this instance."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.Negate()
      Set Negate = .Self
   End With
End Function

'@Description("Converts the string representation of a time interval to its TimeSpan equivalent.")
Public Function Parse(ByVal s As String) As TimeSpan
Attribute Parse.VB_Description = "Converts the string representation of a time interval to its TimeSpan equivalent."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.Parse(s)
      Set Parse = .Self
   End With
End Function

'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified culture-specific format information.")
Public Function Parse2(ByVal strInput As String, ByRef formatProvider As IFormatProvider) As TimeSpan
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.Parse2(strInput, formatProvider)
      Set Parse2 = .Self
   End With
End Function

'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")
Public Function ParseExact(ByVal strInput As String, ByVal format As String, ByRef formatProvider As IFormatProvider) As TimeSpan
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.ParseExact(strInput, format, formatProvider)
      Set ParseExact = .Self
   End With
End Function

'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified array of format strings and culture-specific format information. The format of the string representation must match one of the specified formats exactly.")
Public Function ParseExact2(ByVal strInput As String, ByRef formats() As String, ByRef formatProvider As IFormatProvider) As TimeSpan
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.ParseExact2(strInput, formats, formatProvider)
      Set ParseExact2 = .Self
   End With
End Function

'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format, culture-specific format information, and styles. The format of the string representation must match the specified format exactly.")
Public Function ParseExact3(ByVal strInput As String, ByVal format As String, ByRef formatProvider As IFormatProvider, ByVal styles As DotNetLib.TimeSpanStyles) As TimeSpan
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.ParseExact3(strInput, format, formatProvider, styles)
      Set ParseExact3 = .Self
   End With
End Function

'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified formats, culture-specific format information, and styles. The format of the string representation must match one of the specified formats exactly")
Public Function ParseExact4(ByVal strInput As String, ByRef formats() As String, ByRef formatProvider As IFormatProvider, ByVal styles As DotNetLib.TimeSpanStyles) As TimeSpan
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.ParseExact4(strInput, formats, formatProvider, styles)
      Set ParseExact4 = .Self
   End With
End Function

'@Description("Returns a new TimeSpan object whose value is the difference between the specified TimeSpan object and this instance.")
Public Function Subtract(ByRef ts As TimeSpan) As TimeSpan
Attribute Subtract.VB_Description = "Returns a new TimeSpan object whose value is the difference between the specified TimeSpan object and this instance."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.Subtract(ts.TimeSpanComObject)
      Set Subtract = .Self
   End With
End Function

'@DefaultMember
'@Description("Converts the value of the current TimeSpan object to its equivalent string representation.")
Public Property Get ToString() As String
Attribute ToString.VB_Description = "Converts the value of the current TimeSpan object to its equivalent string representation."
Attribute ToString.VB_UserMemId = 0
   ToString = this.DotNetLibTimeSpan.ToString()
End Property

'@Description("Converts the value of the current TimeSpan object to its equivalent string representation by using the specified format.")
Public Function ToString2(ByVal format As String) As String
Attribute ToString2.VB_Description = "Converts the value of the current TimeSpan object to its equivalent string representation by using the specified format."
   ToString2 = this.DotNetLibTimeSpan.ToString2(format)
End Function

'@Description("Converts the value of the current TimeSpan object to its equivalent string representation by using the specified format and culture-specific formatting information.")
Public Function ToString3(ByVal format As String, ByRef formatProvider As IFormatProvider) As String
Attribute ToString3.VB_Description = "Converts the value of the current TimeSpan object to its equivalent string representation by using the specified format and culture-specific formatting information."
   ToString3 = this.DotNetLibTimeSpan.ToString3(format, formatProvider)
End Function

'@Description("Converts the string representation of a time interval to its TimeSpan equivalent and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse(ByVal s As String, ByRef result As TimeSpan) As Boolean
Attribute TryParse.VB_Description = "Converts the string representation of a time interval to its TimeSpan equivalent and returns a value that indicates whether the conversion succeeded."
   TryParse = this.DotNetLibTimeSpan.TryParse(s, result.TimeSpanComObject)
End Function

'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified culture-specific formatting information, and returns a value that indicates whether the conversion succeeded.")
Public Function TryParse2(ByVal strInput As String, ByRef formatProvider As IFormatProvider, ByRef result As TimeSpan) As Boolean
   TryParse2 = this.DotNetLibTimeSpan.TryParse2(strInput, formatProvider, result.TimeSpanComObject)
End Function

'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")
Public Function TryParseExact(ByVal strInput As String, ByVal format As String, ByRef formatProvider As IFormatProvider, ByRef result As TimeSpan) As Boolean
   TryParseExact = this.DotNetLibTimeSpan.TryParseExact(strInput, format, formatProvider, result.TimeSpanComObject)
End Function

'@Description("Converts the specified string representation of a time interval to its TimeSpan equivalent by using the specified formats and culture-specific format information. The format of the string representation must match one of the specified formats exactly.")
Public Function TryParseExact2(ByVal strInput As String, ByRef formats() As String, ByRef formatProvider As IFormatProvider, ByRef result As TimeSpan) As Boolean
   TryParseExact2 = this.DotNetLibTimeSpan.TryParseExact2(strInput, formats, formatProvider, result.TimeSpanComObject)
End Function

'@Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format, culture-specific format information and styles. The format of the string representation must match the specified format exactly.")
Public Function TryParseExact3(ByVal strInput As String, ByVal format As String, ByRef formatProvider As IFormatProvider, ByVal styles As DotNetLib.TimeSpanStyles, ByRef result As TimeSpan) As Boolean
   TryParseExact3 = this.DotNetLibTimeSpan.TryParseExact3(strInput, format, formatProvider, styles, result.TimeSpanComObject)
End Function

'@Description("Converts the specified string representation of a time interval to its TimeSpan equivalent by using the specified formats, culture-specific format information and styles. The format of the string representation must match one of the specified formats exactly.")
Public Function TryParseExact4(ByVal strInput As String, ByRef formats() As String, ByRef formatProvider As IFormatProvider, ByVal styles As DotNetLib.TimeSpanStyles, ByRef result As TimeSpan) As Boolean
   TryParseExact4 = this.DotNetLibTimeSpan.TryParseExact4(strInput, formats, formatProvider, styles, result.TimeSpanComObject)
End Function

'@Description("Adds two specified TimeSpan instances.")
Public Function Addition(ByRef t1 As TimeSpan, ByRef t2 As TimeSpan) As TimeSpan
Attribute Addition.VB_Description = "Adds two specified TimeSpan instances."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.Addition(t1.TimeSpanComObject, t2.TimeSpanComObject)
      Set Addition = .Self
   End With
End Function

'@Description("Indicates whether two TimeSpan instances are equal.")
Public Function Equality(ByRef t1 As TimeSpan, ByRef t2 As TimeSpan) As Boolean
Attribute Equality.VB_Description = "Indicates whether two TimeSpan instances are equal."
   Equality = this.DotNetLibTimeSpan.Equality(t1.TimeSpanComObject, t2.TimeSpanComObject)
End Function

'@Description("Indicates whether a specified TimeSpan is greater than another specified TimeSpan.")
Public Function GreaterThan(ByRef t1 As TimeSpan, ByRef t2 As TimeSpan) As Boolean
Attribute GreaterThan.VB_Description = "Indicates whether a specified TimeSpan is greater than another specified TimeSpan."
   GreaterThan = this.DotNetLibTimeSpan.GreaterThan(t1.TimeSpanComObject, t2.TimeSpanComObject)
End Function

'@Description("Indicates whether a specified TimeSpan is greater than or equal to another specified TimeSpan.")
Public Function GreaterThanOrEqual(ByRef t1 As TimeSpan, ByRef t2 As TimeSpan) As Boolean
Attribute GreaterThanOrEqual.VB_Description = "Indicates whether a specified TimeSpan is greater than or equal to another specified TimeSpan."
   GreaterThanOrEqual = this.DotNetLibTimeSpan.GreaterThanOrEqual(t1.TimeSpanComObject, t2.TimeSpanComObject)
End Function

'@Description("Indicates whether two TimeSpan instances are not equal.")
Public Function Inequality(ByRef t1 As TimeSpan, ByRef t2 As TimeSpan) As Boolean
Attribute Inequality.VB_Description = "Indicates whether two TimeSpan instances are not equal."
   Inequality = this.DotNetLibTimeSpan.Inequality(t1.TimeSpanComObject, t2.TimeSpanComObject)
End Function

'@Description("Indicates whether a specified TimeSpan is less than another specified TimeSpan.")
Public Function LessThan(ByRef t1 As TimeSpan, ByRef t2 As TimeSpan) As Boolean
Attribute LessThan.VB_Description = "Indicates whether a specified TimeSpan is less than another specified TimeSpan."
   LessThan = this.DotNetLibTimeSpan.LessThan(t1.TimeSpanComObject, t2.TimeSpanComObject)
End Function

'@Description("Indicates whether a specified TimeSpan is less than or equal to another specified TimeSpan.")
Public Function LessThanOrEqual(ByRef t1 As TimeSpan, ByRef t2 As TimeSpan) As Boolean
Attribute LessThanOrEqual.VB_Description = "Indicates whether a specified TimeSpan is less than or equal to another specified TimeSpan."
   LessThanOrEqual = this.DotNetLibTimeSpan.LessThanOrEqual(t1.TimeSpanComObject, t2.TimeSpanComObject)
End Function

'@Description("Subtracts a specified TimeSpan from another specified TimeSpan.")
Public Function Subtraction(ByRef t1 As TimeSpan, ByRef t2 As TimeSpan) As TimeSpan
Attribute Subtraction.VB_Description = "Subtracts a specified TimeSpan from another specified TimeSpan."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.Subtraction(t1.TimeSpanComObject, t2.TimeSpanComObject)
      Set Subtraction = .Self
   End With
End Function

'@Description("Returns a TimeSpan whose value is the negated value of the specified instance.")
Public Function UnaryNegation(ByRef t As TimeSpan) As TimeSpan
Attribute UnaryNegation.VB_Description = "Returns a TimeSpan whose value is the negated value of the specified instance."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.UnaryNegation(t.TimeSpanComObject)
      Set UnaryNegation = .Self
   End With
End Function

'@Description("Returns the specified instance of TimeSpan.")
Public Function UnaryPlus(ByRef t As TimeSpan) As TimeSpan
Attribute UnaryPlus.VB_Description = "Returns the specified instance of TimeSpan."
   With New TimeSpan
      Set .TimeSpanComObject = this.DotNetLibTimeSpan.UnaryPlus(t.TimeSpanComObject)
      Set UnaryPlus = .Self
   End With
End Function
