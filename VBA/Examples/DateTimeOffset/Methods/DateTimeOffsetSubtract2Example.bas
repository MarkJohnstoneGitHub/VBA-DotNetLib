Attribute VB_Name = "DateTimeOffsetSubtract2Example"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified July 21, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.subtract?view=netframework-4.8.1#system-datetimeoffset-subtract(system-timespan)

Option Explicit

'@Description("The following example illustrates subtraction that uses the Subtract method.")
Public Sub DateTimeOffsetSubtract2()
Attribute DateTimeOffsetSubtract2.VB_Description = "The following example illustrates subtraction that uses the Subtract method."
    Dim offsetDate As DateTimeOffset
    Set offsetDate = DateTimeOffset.CreateFromDateTimeParts(2007, 12, 3, 11, 30, 0, TimeSpan.Create(-8, 0, 0))
    Dim duration As TimeSpan
    Set duration = TimeSpan.Create2(7, 18, 0, 0)
    Debug.Print offsetDate.Subtract2(duration).ToString()
  
' Output:
' 11/25/2007 5:30:00 PM -08:00
End Sub
