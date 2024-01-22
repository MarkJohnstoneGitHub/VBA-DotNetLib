Attribute VB_Name = "DateTimeTryParseExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 14, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tryparse?view=netframework-4.8.1#system-datetime-tryparse(system-string-system-datetime@)

Option Explicit

'@Description("The following example passes a number of date and time strings to the DateTime.TryParse(String, DateTime) method.")
Public Sub DateTimeTryParse()
Attribute DateTimeTryParse.VB_Description = "The following example passes a number of date and time strings to the DateTime.TryParse(String, DateTime) method."
   Dim dateStrings() As String
   dateStrings = StringArray.CreateInitialize1D("05/01/2009 14:57:32.8", "2009-05-01 14:57:32.8", _
                                 "2009-05-01T14:57:32.8375298-04:00", "5/01/2008", _
                                 "5/01/2008 14:57:32.80 -07:00", _
                                 "1 May 2008 2:57:32.8 PM", "16-05-2009 1:00:32 PM", _
                                 "Fri, 15 May 2009 20:10:57 GMT")
   
    Debug.Print VBString.Format("Attempting to parse strings using {0} culture.", _
                               CultureInfo.CurrentCulture.name)
    Dim dateValue As DotNetLib.DateTime
    Dim dateString As Variant
    For Each dateString In dateStrings
        If (DateTime.TryParse(dateString, dateValue)) Then
            Debug.Print VBString.Format("  Converted '{0}' to {1} ({2}).", dateString, _
                              dateValue, DateTimeKindHelper.ToString(dateValue.Kind))
        Else
            Debug.Print VBString.Format("  Unable to parse '{0}'.", dateString)
        End If
    Next
End Sub

' The example displays output like the following:
'    Attempting to parse strings using en-US culture.
'      Converted '05/01/2009 14:57:32.8' to 5/1/2009 2:57:32 PM (Unspecified).
'      Converted '2009-05-01 14:57:32.8' to 5/1/2009 2:57:32 PM (Unspecified).
'      Converted '2009-05-01T14:57:32.8375298-04:00' to 5/1/2009 11:57:32 AM (Local).
'
'      Converted '5/01/2008' to 5/1/2008 12:00:00 AM (Unspecified).
'      Converted '5/01/2008 14:57:32.80 -07:00' to 5/1/2008 2:57:32 PM (Local).
'      Converted '1 May 2008 2:57:32.8 PM' to 5/1/2008 2:57:32 PM (Unspecified).
'      Unable to parse '16-05-2009 1:00:32 PM'.
'      Converted 'Fri, 15 May 2009 20:10:57 GMT' to 5/15/2009 1:10:57 PM (Local).
