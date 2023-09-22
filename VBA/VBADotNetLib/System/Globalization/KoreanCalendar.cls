VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KoreanCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents the Korean calendar."
'@Folder "VBADotNetLib.System.Globalization"
'@PredeclaredId
'@ModuleDescription("Represents the Korean calendar.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 19, 2023
'@LastModified September 22, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.koreancalendar?view=netframework-4.8.1
'
'@Remarks
' Singleton class for KoreanCalendar containing factory methods and static members.

Option Explicit

Private Type TKoreanCalendar
   KoreanCalendar As DotNetLib.KoreanCalendarSingleton
End Type

Private this As TKoreanCalendar

Private Sub Class_Initialize()
   Set this.KoreanCalendar = New DotNetLib.KoreanCalendarSingleton
End Sub

Private Sub Class_Terminate()
   Set this.KoreanCalendar = Nothing
End Sub

'@Static
'@Description("Initializes a new instance of the HebrewCalendar class.")
Public Function Create() As DotNetLib.KoreanCalendar
Attribute Create.VB_Description = "Initializes a new instance of the HebrewCalendar class."
   Set Create = this.KoreanCalendar.Create()
End Function

'@Static
'@Description("Represents the current era of the current calendar. The value of this field is 0.")
Public Property Get CurrentEra() As Long
Attribute CurrentEra.VB_Description = "Represents the current era of the current calendar. The value of this field is 0."
   CurrentEra = this.KoreanCalendar.CurrentEra()
End Property

'@Static
'@Description("Represents the current era. This field is constant.")
Public Property Get KoreanEra() As Long
Attribute KoreanEra.VB_Description = "Represents the current era. This field is constant."
   KoreanEra = this.KoreanCalendar.KoreanEra()
End Property

'@Static
'@Description("Determines whether the specified object instances are considered equal.")
Public Function Equals(ByVal objA As Variant, ByVal objB As Variant) As Boolean
Attribute Equals.VB_Description = "Determines whether the specified object instances are considered equal."
   Equals = this.KoreanCalendar.Equals(objA, objB)
End Function

'@Static
'@Description("Returns a read-only version of the specified Calendar object.")
Public Function ReadOnly(ByVal Calendar As DotNetLib.Calendar) As DotNetLib.Calendar
Attribute ReadOnly.VB_Description = "Returns a read-only version of the specified Calendar object."
   Set ReadOnly = this.KoreanCalendar.ReadOnly(Calendar)
End Function
