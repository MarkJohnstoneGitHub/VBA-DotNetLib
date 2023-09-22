VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HijriCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents the Hijri calendar."
'@Folder "VBADotNetLib.System.Globalization"
'@PredeclaredId
'@ModuleDescription("Represents the Hijri calendar.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 19, 2023
'@LastModified September 22, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.hijricalendar?view=netframework-4.8.1

'@Remarks
' Singleton class for HijriCalendar containing factory methods and static members.

Option Explicit

Private Type THijriCalendar
   HijriCalendar As DotNetLib.HijriCalendarSingleton
End Type

Private this As THijriCalendar

Private Sub Class_Initialize()
   Set this.HijriCalendar = New DotNetLib.HijriCalendarSingleton
End Sub

Private Sub Class_Terminate()
   Set this.HijriCalendar = Nothing
End Sub

'@Static
'@Description("Initializes a new instance of the HijriCalendar class.")
Public Function Create() As DotNetLib.HijriCalendar
Attribute Create.VB_Description = "Initializes a new instance of the HijriCalendar class."
   Set Create = this.HijriCalendar.Create()
End Function

'@Static
'@Description("Represents the current era of the current calendar. The value of this field is 0.")
Public Property Get CurrentEra() As Long
Attribute CurrentEra.VB_Description = "Represents the current era of the current calendar. The value of this field is 0."
   CurrentEra = this.HijriCalendar.CurrentEra()
End Property

'@Static
'@Description("Represents the current era. This field is constant.")
Public Property Get HijriEra() As Long
Attribute HijriEra.VB_Description = "Represents the current era. This field is constant."
   HijriEra = this.HijriCalendar.HijriEra()
End Property

'@Static
'@Description("Determines whether the specified object instances are considered equal.")
Public Function Equals(ByVal objA As Variant, ByVal objB As Variant) As Boolean
Attribute Equals.VB_Description = "Determines whether the specified object instances are considered equal."
   Equals = this.HijriCalendar.Equals(objA, objB)
End Function

'@Static
'@Description("Returns a read-only version of the specified Calendar object.")
Public Function ReadOnly(ByVal Calendar As DotNetLib.Calendar) As DotNetLib.Calendar
Attribute ReadOnly.VB_Description = "Returns a read-only version of the specified Calendar object."
   Set ReadOnly = this.HijriCalendar.ReadOnly(Calendar)
End Function
