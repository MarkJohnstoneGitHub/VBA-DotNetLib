Attribute VB_Name = "DateTimeDayOfWeekExample"
'Rubberduck annotations
'@Folder "Examples.System.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 9, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.dayofweek?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates the DayOfWeek property and the DayOfWeek enumeration."
Public Sub DateTimeDayOfWeek()
    ' Assume the current culture is en-US.
    ' Create a DateTime for the first of May, 2003.
    Dim dt As DotNetLib.DateTime
    Set dt = DateTime.CreateFromDate(2003, 5, 1)
    Debug.Print VBString.Format("Is Thursday the day of the week for {0:d}?: {1}", _
                       dt, IIf(dt.DayOfWeek = DayOfWeek.DayOfWeek_Thursday, True, False))
    Debug.Print VBString.Format("The day of the week for {0:d} is {1}.", dt, DayOfWeekHelper.ToString(dt.DayOfWeek))
End Sub

'This example produces the following results:
'
'Is Thursday the day of the week for 5/1/2003?: True
'The day of the week for 5/1/2003 is Thursday.


