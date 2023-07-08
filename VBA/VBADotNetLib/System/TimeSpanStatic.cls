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
    
   With New TimeSpan
      Set .ComObject = TimeSpan.ComObject.Zero
      Set this.Zero = .Self
   End With
End Sub

'@StaticField
'@Description("Represents the maximum TimeSpan value. This field is read-only.")
Public Property Get MaxValue() As TimeSpan
   Set MaxValue = this.MaxValue
End Property

'@StaticField
'@Description("Represents the minimum TimeSpan value. This field is read-only.")
Public Property Get MinValue() As TimeSpan
   Set MinValue = this.MinValue
End Property

'@StaticField
'@Description("Represents the number of ticks in 1 day. This field is constant.")
Public Property Get TicksPerDay() As LongLong
   TicksPerDay = this.TicksPerDay
End Property

'@StaticField
'@Description("Represents the number of ticks in 1 hour. This field is constant.")
Public Property Get TicksPerHour() As LongLong
   TicksPerHour = this.TicksPerHour
End Property

'@StaticField
'@Description("Represents the number of ticks in 1 millisecond. This field is constant.")
Public Property Get TicksPerMillisecond() As LongLong
   TicksPerMillisecond = this.TicksPerMillisecond
End Property

'@StaticField
'@Description("Represents the number of ticks in 1 minute. This field is constant.")
Public Property Get TicksPerMinute() As LongLong
   TicksPerMinute = this.TicksPerMinute
End Property

'@Description("Represents the zero TimeSpan value. This field is read-only.")
Public Property Get Zero() As TimeSpan
   Set Zero = this.Zero
End Property
