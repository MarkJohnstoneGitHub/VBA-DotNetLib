Attribute VB_Name = "DateTimeOffsetNowExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.now?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the Now property to retrieve the current date and time and displays it by using each of the standard date and time format strings supported by the DateTimeOffset type.")
Public Sub DateTimeOffsetNow()
Attribute DateTimeOffsetNow.VB_Description = "The following example uses the Now property to retrieve the current date and time and displays it by using each of the standard date and time format strings supported by the DateTimeOffset type."
    Dim fmtStrings() As String
    fmtStrings = Strings.ToArray("d", "D", "f", "F", "g", "G", "M", _
                                  "R", "s", "t", "T", "u", "y")
    
    Dim value As IDateTimeOffset
    Set value = DateTimeOffset.Now
    ' Display date in default format.
    Debug.Print value.ToString()
    Debug.Print
    
    ' Display date using each of the specified formats.
    Dim fmtString As Variant
    For Each fmtString In fmtStrings
       Debug.Print fmtString & " --> " & value.ToString2(fmtString)
    Next
End Sub

' The example displays output similar to the following:
'    11/19/2012 10:57:11 AM -08:00
'
'    d --> 11/19/2012
'    D --> Monday, November 19, 2012
'    f --> Monday, November 19, 2012 10:57 AM
'    F --> Monday, November 19, 2012 10:57:11 AM
'    g --> 11/19/2012 10:57 AM
'    G --> 11/19/2012 10:57:11 AM
'    M --> November 19
'    R --> Mon, 19 Nov 2012 18:57:11 GMT
'    s --> 2012-11-19T10:57:11
'    t --> 10:57 AM
'    T --> 10:57:11 AM
'    u --> 2012-11-19 18:57:11Z
'    y --> November, 2012
