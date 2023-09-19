VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Calendar static members."
'@Folder "VBADotNetLib.System.Globalization"
'@PredeclaredId
'@ModuleDescription("Calendar static members.")

'@Singleton
'@Abstract

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 19, 2023
'@LastModified September 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.calendar?view=netframework-4.8.1

'@Remarks
' Singleton class for Calendar containing static members.
' Calendar is an abstract class so therefore has no factory methods.

Option Explicit

Private Type TCalendarSingleton
   CalendarSingleton As DotNetLib.CalendarSingleton
End Type

Private this As TCalendarSingleton

Private Sub Class_Initialize()
   Set this.CalendarSingleton = New DotNetLib.CalendarSingleton
End Sub

Private Sub Class_Terminate()
   Set this.CalendarSingleton = Nothing
End Sub

'@Description("Represents the current era of the current calendar. The value of this field is 0.")
Public Property Get CurrentEra() As Long
Attribute CurrentEra.VB_Description = "Represents the current era of the current calendar. The value of this field is 0."
   CurrentEra = this.CalendarSingleton.CurrentEra()
End Property

'@Description("Determines whether the specified object instances are considered equal.")
Public Function Equals(ByVal objA As Variant, ByVal objB As Variant) As Boolean
Attribute Equals.VB_Description = "Determines whether the specified object instances are considered equal."
   Equals = this.CalendarSingleton.Equals(objA, objB)
End Function

'@Description("Returns a read-only version of the specified Calendar object.")
Public Function ReadOnly(ByRef calendar As DotNetLib.ICalendar) As ICalendar
Attribute ReadOnly.VB_Description = "Returns a read-only version of the specified Calendar object."
   Set ReadOnly = this.CalendarSingleton.ReadOnly(calendar)
End Function
