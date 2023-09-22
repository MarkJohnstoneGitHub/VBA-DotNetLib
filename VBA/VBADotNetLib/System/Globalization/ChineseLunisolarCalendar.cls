VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChineseLunisolarCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents time in divisions, such as months, days, and years. Years are calculated using the Chinese calendar, while days and months are calculated using the lunisolar calendar."
'@Folder "VBADotNetLib.System.Globalization"
'@PredeclaredId
'@ModuleDescription("Represents time in divisions, such as months, days, and years. Years are calculated using the Chinese calendar, while days and months are calculated using the lunisolar calendar.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 19, 2023
'@LastModified September 22, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.chineselunisolarcalendar?view=netframework-4.8.1

'@Remarks
' Singleton class for ChineseLunisolarCalendar containing factory methods and static members.

Option Explicit

Private Type TChineseLunisolarCalendar
   ChineseLunisolarCalendar As DotNetLib.ChineseLunisolarCalendarSingleton
End Type

Private this As TChineseLunisolarCalendar

Private Sub Class_Initialize()
   Set this.ChineseLunisolarCalendar = New DotNetLib.ChineseLunisolarCalendarSingleton
End Sub

Private Sub Class_Terminate()
   Set this.ChineseLunisolarCalendar = Nothing
End Sub

'@Static
'@Description("Initializes a new instance of the ChineseLunisolarCalendar class.")
Public Function Create() As DotNetLib.ChineseLunisolarCalendar
Attribute Create.VB_Description = "Initializes a new instance of the ChineseLunisolarCalendar class."
   Set Create = this.ChineseLunisolarCalendar.Create()
End Function

'@Static
'@Description("Represents the current era of the current calendar. The value of this field is 0.")
Public Property Get CurrentEra() As Long
Attribute CurrentEra.VB_Description = "Represents the current era of the current calendar. The value of this field is 0."
   CurrentEra = this.ChineseLunisolarCalendar.CurrentEra()
End Property

'@Static
'@Description("Specifies the era that corresponds to the current ChineseLunisolarCalendar object.")
Public Property Get ChineseEra() As Long
Attribute ChineseEra.VB_Description = "Specifies the era that corresponds to the current ChineseLunisolarCalendar object."
   ChineseEra = this.ChineseLunisolarCalendar.ChineseEra()
End Property

'@Static
'@Description("Determines whether the specified object instances are considered equal.")
Public Function Equals(ByVal objA As Variant, ByVal objB As Variant) As Boolean
Attribute Equals.VB_Description = "Determines whether the specified object instances are considered equal."
   Equals = this.ChineseLunisolarCalendar.Equals(objA, objB)
End Function

'@Static
'@Description("Returns a read-only version of the specified Calendar object.")
Public Function ReadOnly(ByVal Calendar As DotNetLib.Calendar) As DotNetLib.Calendar
Attribute ReadOnly.VB_Description = "Returns a read-only version of the specified Calendar object."
   Set ReadOnly = this.ChineseLunisolarCalendar.ReadOnly(Calendar)
End Function
