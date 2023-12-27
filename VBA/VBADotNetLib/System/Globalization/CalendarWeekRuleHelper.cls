VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarWeekRuleHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBADotNetLib.System.Globalization")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 27, 2023
'@LastModified December 27, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.calendarweekrule?view=netframework-4.8.1

Option Explicit

'@DefaultMember
'@Description("Converts the value of a CalendarWeekRule to string value.")
Public Property Get ToString(ByVal value As mscorlib.CalendarWeekRule) As String
    Select Case value
        Case CalendarWeekRule_FirstDay: ToString = "FirstDay"
        Case CalendarWeekRule_FirstFourDayWeek: ToString = "FirstFourDayWeek"
        Case CalendarWeekRule_FirstFullWeek: ToString = "FirstFullWeek"
    End Select
End Property
