Attribute VB_Name = "DateTimeOffsetAddYearsExample"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 20, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.addyears?view=netframework-4.8.1#examples

Option Explicit

''
' In the United States, driver's licenses cannot be issued to persons under 16 years of age.
' The following example displays the latest possible date on which a person must be born in
' order to legally be issued a driver's license.
''
Public Sub DateTimeOffsetAddYears()
    Const minimumAge As Long = 16
    Dim dateToday As DotNetLib.DateTimeOffset
    Set dateToday = DateTimeOffset.Now
    Dim latestBirthday As DotNetLib.DateTimeOffset
    Set latestBirthday = dateToday.AddYears(-1 * minimumAge)
    Debug.Print VBString.Format("To possess a driver's license, you must have been born on or before {0:d}.", _
                  latestBirthday)
End Sub

' Output:
' To possess a driver's license, you must have been born on or before 20/07/2007.


