VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThaiBuddhistCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "VBADotNetLib.System.Globalization"
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 19, 2023
'@LastModified September 22, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.thaibuddhistcalendar?view=netframework-4.8.1

'@Remarks
' Singleton class for ThaiBuddhistCalendar containing factory methods and static members.

Option Explicit

Private Type TThaiBuddhistCalendar
   ThaiBuddhistCalendarSingleton As DotNetLib.ThaiBuddhistCalendarSingleton
End Type

Private this As TThaiBuddhistCalendar

Private Sub Class_Initialize()
   Set this.ThaiBuddhistCalendarSingleton = New DotNetLib.ThaiBuddhistCalendarSingleton
End Sub

Private Sub Class_Terminate()
   Set this.ThaiBuddhistCalendarSingleton = Nothing
End Sub

'@Static
'@Description("Initializes a new instance of the ThaiBuddhistCalendar class.")
Public Function Create() As DotNetLib.ThaiBuddhistCalendar
Attribute Create.VB_Description = "Initializes a new instance of the ThaiBuddhistCalendar class."
   Set Create = this.ThaiBuddhistCalendarSingleton.Create()
End Function

'@Static
'@Description("Represents the current era of the current calendar. The value of this field is 0.")
Public Property Get CurrentEra() As Long
Attribute CurrentEra.VB_Description = "Represents the current era of the current calendar. The value of this field is 0."
   CurrentEra = this.ThaiBuddhistCalendarSingleton.CurrentEra()
End Property

'@Static
'@Description("Represents the current era. This field is constant.")
Public Property Get ThaiBuddhistEra() As Long
Attribute ThaiBuddhistEra.VB_Description = "Represents the current era. This field is constant."
   ThaiBuddhistEra = this.ThaiBuddhistCalendarSingleton.ThaiBuddhistEra()
End Property

'@Static
'@Description("Determines whether the specified object instances are considered equal.")
Public Function Equals(ByVal objA As Variant, ByVal objB As Variant) As Boolean
Attribute Equals.VB_Description = "Determines whether the specified object instances are considered equal."
   Equals = this.ThaiBuddhistCalendarSingleton.Equals(objA, objB)
End Function

'@Static
'@Description("Returns a read-only version of the specified Calendar object.")
Public Function ReadOnly(ByVal Calendar As DotNetLib.Calendar) As DotNetLib.Calendar
Attribute ReadOnly.VB_Description = "Returns a read-only version of the specified Calendar object."
   Set ReadOnly = this.ThaiBuddhistCalendarSingleton.ReadOnly(Calendar)
End Function
