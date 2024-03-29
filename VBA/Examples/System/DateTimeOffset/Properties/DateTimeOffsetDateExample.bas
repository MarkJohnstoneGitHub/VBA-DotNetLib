Attribute VB_Name = "DateTimeOffsetDateExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 18, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.date?view=netframework-4.8.1#examples

Option Explicit

''
' The following example retrieves the value of the Date property for a specific date.
' It then displays that value to the console using some standard and custom date-only
' format specifiers.
''
Public Sub DateTimeOffsetDate()
    ' Illustrate Date property and date formatting
    Dim thisDate As DotNetLib.DateTimeOffset
    Set thisDate = DateTimeOffset.CreateFromDateTimeParts(2008, 3, 17, 1, 32, 0, TimeSpan.Create(-5, 0, 0))
    Dim fmt As String    ' format specifier
    
    ' Display date only using "D" format specifier
    ' For en-us culture, displays:
    '   'D' format specifier: Monday, March 17, 2008
    fmt = "D"
    Debug.Print VBString.Format("'{0}' format specifier: {1}", _
                                fmt, thisDate.Date.ToString2(fmt))
    
    ' Display date only using "d" format specifier
    ' For en-us culture, displays:
    '   'd' format specifier: 3/17/2008
    fmt = "d"
    Debug.Print VBString.Format("'{0}' format specifier: {1}", _
                                fmt, thisDate.Date.ToString2(fmt))
    
    ' Display date only using "Y" (or "y") format specifier
    ' For en-us culture, displays:
    '   'Y' format specifier: March, 2008
    fmt = "Y"
    Debug.Print VBString.Format("'{0}' format specifier: {1}", _
                                fmt, thisDate.Date.ToString2(fmt))
    
    ' Display date only using custom format specifier
    ' For en-us culture, displays:
    '   'dd MMM yyyy' format specifier: 17 Mar 2008
    fmt = "dd MMM yyyy"
    Debug.Print VBString.Format("'{0}' format specifier: {1}", _
                                fmt, thisDate.Date.ToString2(fmt))
End Sub

' Output:
'   'D' format specifier: Monday, 17 March 2008
'   'd' format specifier: 17/03/2008
'   'Y' format specifier: March 2008
'   'dd MMM yyyy' format specifier: 17 Mar 2008

