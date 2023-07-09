VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeSpanStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rubberduck annotations
'@Folder "VBADotNetLib.System"
'@PredeclaredId

'@Static

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified July 09, 2023

Option Explicit

Private Type TTimeSpanStatic
   MaxValue As TimeSpan
   MinValue As TimeSpan
   TicksPerDay As LongLong
   TicksPerHour As LongLong
   TicksPerMillisecond As LongLong
   TicksPerMinute As LongLong
   TicksPerSecond As LongLong
   Zero As TimeSpan
End Type

Private this As TTimeSpanStatic

Private Sub Class_Initialize()
   With New TimeSpan
      Set .ComObject = TimeSpan.ComObject.MaxValue
      Set this.MaxValue = .Self
   End With
    
   With New TimeSpan
      Set .ComObject = TimeSpan.ComObject.MinValue
      Set this.MinValue = .Self
   End With
    
   this.TicksPerDay = TimeSpan.ComObject.TicksPerDay()
   this.TicksPerHour = TimeSpan.ComObject.TicksPerHour()
   this.TicksPerMillisecond = TimeSpan.ComObject.TicksPerMillisecond()
   this.TicksPerMinute = TimeSpan.ComObject.TicksPerMinute()
   this.TicksPerSecond = TimeSpan.ComObject.TicksPerSecond()
    
   With New TimeSpan
      Set .ComObject = TimeSpan.ComObject.Zero
      Set this.Zero = .Self
   End With
End Sub

'@Static
'@Description("Represents the maximum TimeSpan value. This field is read-only.")
Public Property Get MaxValue() As TimeSpan
   Set MaxValue = this.MaxValue
End Property

'@Static
'@Description("Represents the minimum TimeSpan value. This field is read-only.")
Public Property Get MinValue() As TimeSpan
   Set MinValue = this.MinValue
End Property

'@Static
'@Description("Represents the number of ticks in 1 day. This field is constant.")
Public Property Get TicksPerDay() As LongLong
   TicksPerDay = this.TicksPerDay
End Property

'@Static
'@Description("Represents the number of ticks in 1 hour. This field is constant.")
Public Property Get TicksPerHour() As LongLong
   TicksPerHour = this.TicksPerHour
End Property

'@Static
'@Description("Represents the number of ticks in 1 millisecond. This field is constant.")
Public Property Get TicksPerMillisecond() As LongLong
   TicksPerMillisecond = this.TicksPerMillisecond
End Property

'@Static
'@Description("Represents the number of ticks in 1 minute. This field is constant.")
Public Property Get TicksPerMinute() As LongLong
   TicksPerMinute = this.TicksPerMinute
End Property

'@Static
'@Description("Represents the number of ticks in 1 second.")
Public Property Get TicksPerSecond() As LongLong
   TicksPerSecond = this.TicksPerSecond
End Property

'@Static
'@Description("Represents the zero TimeSpan value. This field is read-only.")
Public Property Get Zero() As TimeSpan
   Set Zero = this.Zero
End Property
