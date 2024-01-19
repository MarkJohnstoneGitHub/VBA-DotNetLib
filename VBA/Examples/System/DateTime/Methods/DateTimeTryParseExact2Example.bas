Attribute VB_Name = "DateTimeTryParseExact2Example"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 15, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tryparseexact?view=netframework-4.8.1#system-datetime-tryparseexact(system-string-system-string()-system-iformatprovider-system-globalization-datetimestyles-system-datetime@)

Option Explicit

' The following example uses the DateTime.TryParseExact(String, String, IFormatProvider, DateTimeStyles, DateTime)
' method to ensure that a string in a number of possible formats can be successfully parsed .
Public Sub DateTimeTryParseExact2()
    Dim formats() As String
    formats = StringArray.CreateInitialize1D("M/d/yyyy h:mm:ss tt", "M/d/yyyy h:mm tt", _
                            "MM/dd/yyyy hh:mm:ss", "M/d/yyyy h:mm:ss", _
                            "M/d/yyyy hh:mm tt", "M/d/yyyy hh tt", _
                            "M/d/yyyy h:mm", "M/d/yyyy h:mm", _
                            "MM/dd/yyyy hh:mm", "M/dd/yyyy hh:mm")
                            
    Dim dateStrings() As String
    dateStrings = StringArray.CreateInitialize1D("5/1/2009 6:32 PM", "05/01/2009 6:32:05 PM", _
                                "5/1/2009 6:32:00", "05/01/2009 06:32", _
                                "05/01/2009 06:32:00 PM", "05/01/2009 06:32:00")

    Dim dateValue As DotNetLib.DateTime
    Dim dateString As Variant
    For Each dateString In dateStrings
        If (DateTime.TryParseExact2(dateString, formats, _
                                    CultureInfo.CreateFromName("en-US"), _
                                    DateTimeStyles.DateTimeStyles_None, _
                                    dateValue)) Then
            Debug.Print VBString.Format("Converted '{0}' to {1}.", dateString, dateValue)
        Else
            Debug.Print VBString.Format("Unable to convert '{0}' to a date.", dateString)
        End If
    Next
End Sub

' The example displays the following output:
'       Converted '5/1/2009 6:32 PM' to 5/1/2009 6:32:00 PM.
'       Converted '05/01/2009 6:32:05 PM' to 5/1/2009 6:32:05 PM.
'       Converted '5/1/2009 6:32:00' to 5/1/2009 6:32:00 AM.
'       Converted '05/01/2009 06:32' to 5/1/2009 6:32:00 AM.
'       Converted '05/01/2009 06:32:00 PM' to 5/1/2009 6:32:00 PM.
'       Converted '05/01/2009 06:32:00' to 5/1/2009 6:32:00 AM.


