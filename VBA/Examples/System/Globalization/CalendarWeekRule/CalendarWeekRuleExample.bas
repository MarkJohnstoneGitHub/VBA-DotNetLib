Attribute VB_Name = "CalendarWeekRuleExample"
'@Folder "Examples.System.Globalization.CalendarWeekRule"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 27, 2023
'@LastModified December 27, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.calendarweekrule?view=netframework-4.8.1

Option Explicit

Private cal As DotNetLib.Calendar

Public Sub CalendarWeekRuleExample()
    Set cal = GregorianCalendar.Create()
        
    Dim pvtDate As DotNetLib.DateTime
    Set pvtDate = DateTime.CreateFromDate(2013, 1, 5)
    Dim firstDay As mscorlib.DayOfWeek
    firstDay = DayOfWeek.DayOfWeek_Sunday
    Dim pvtRule As mscorlib.CalendarWeekRule
    
    pvtRule = CalendarWeekRule.CalendarWeekRule_FirstFullWeek
    Call ShowWeekNumber(pvtDate, pvtRule, firstDay)
    
    pvtRule = CalendarWeekRule.CalendarWeekRule_FirstFourDayWeek
    Call ShowWeekNumber(pvtDate, pvtRule, firstDay)

    Debug.Print
    Set pvtDate = DateTime.CreateFromDate(2010, 1, 3)
    Call ShowWeekNumber(pvtDate, pvtRule, firstDay)
End Sub

Private Sub ShowWeekNumber(ByVal dat As DotNetLib.DateTime, ByVal pRule As mscorlib.CalendarWeekRule, ByVal firstDay As DotNetLib.DayOfWeek)
    Debug.Print VBString.Format("{0:d} with {1} rule and {2} as first day of week: week {3}", _
                        dat, CalendarWeekRuleHelper.ToString(pRule), DayOfWeekHelper.ToString(firstDay), cal.GetWeekOfYear(dat, pRule, firstDay))
End Sub

' The example displays the following output:
'       1/5/2013 with FirstFullWeek rule and Sunday as first day of week: week 53
'       1/5/2013 with FirstFourDayWeek rule and Sunday as first day of week: week 1
'
'       1/3/2010 with FirstFourDayWeek rule and Sunday as first day of week: week 1


