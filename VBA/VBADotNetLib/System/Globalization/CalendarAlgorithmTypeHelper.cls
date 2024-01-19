VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarAlgorithmTypeHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "VBADotNetLib.System.Globalization"
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 27, 2023
'@LastModified December 27, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.calendaralgorithmtype?view=netframework-4.8.1

Option Explicit

'@DefaultMember
'@Description("Converts the value of a CalendarWeekRule to string value.")
Public Property Get ToString(ByVal value As mscorlib.CalendarAlgorithmType) As String
Attribute ToString.VB_Description = "Converts the value of a CalendarWeekRule to string value."
Attribute ToString.VB_UserMemId = 0
    Select Case value
        Case CalendarAlgorithmType_LunarCalendar: ToString = "LunarCalendar"
        Case CalendarAlgorithmType_LunisolarCalendar: ToString = "LunisolarCalendar"
        Case CalendarAlgorithmType_SolarCalendar: ToString = "SolarCalendar"
        Case CalendarAlgorithmType_Unknown: ToString = "Unknown"
    End Select
End Property
