Attribute VB_Name = "CalendarAlgorithmTypeExample"
'@Folder "Examples.System.Globalization.CalendarAlgorithmType"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 27, 2023
'@LastModified December 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.calendaralgorithmtype?view=netframework-4.8.1#examples

Option Explicit

''
' This example demonstrates the Calendar.AlgorithmType property and
' CalendarAlgorithmType enumeration.
''
Public Sub CalendarAlgorithmTypeExample()
    Dim grCal As DotNetLib.GregorianCalendar
    Set grCal = GregorianCalendar.Create()
    Dim hiCal As DotNetLib.HijriCalendar
    Set hiCal = HijriCalendar.Create()
    
    '@TODO Implement JapaneseLunisolarCalendar
    'dim jaCal as DotNetLib.JapaneseLunisolarCalendar
    'Set jaCal = JapaneseLunisolarCalendar.Create()
    
    Call Display(grCal)
    Call Display(hiCal)
    'Call Display(jaCal)
End Sub

Private Sub Display(ByVal cal As DotNetLib.Calendar)
    Dim pvtName As DotNetLib.String
    Set pvtName = Strings.Copy(cal.ToString()).PadRight(50, ".")
    Debug.Print VBString.Format("{0} {1}", pvtName, CalendarAlgorithmTypeHelper.ToString(cal.AlgorithmType))
End Sub

'/*
'This code example produces the following results:
'
'System.Globalization.GregorianCalendar............ SolarCalendar
'System.Globalization.HijriCalendar................ LunarCalendar
'System.Globalization.JapaneseLunisolarCalendar.... LunisolarCalendar
'
'*/
