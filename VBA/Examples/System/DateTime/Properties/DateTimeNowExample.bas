Attribute VB_Name = "DateTimeNowExample"
'@Folder "Examples.System.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 14, 2023
'@LastModified August 14, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.now?view=netframework-4.8.1#examples

Option Explicit

Public Sub DateTimeNow()
    Dim localDate As DotNetLib.DateTime
    Set localDate = DateTime.Now
    Dim utcDate As DotNetLib.DateTime
    Set utcDate = DateTime.UtcNow
    
    Dim cultureNames() As String
    cultureNames = Strings.ToArray("en-US", "en-GB", "fr-FR", _
                                    "de-DE", "ru-RU")
    Dim cultureName As Variant
    For Each cultureName In cultureNames
        Dim culture As DotNetLib.CultureInfo
        Set culture = CultureInfo.Create2(cultureName)
        Debug.Print culture.NativeName; ":"
        Debug.Print "   Local date and time: "; localDate.ToString3(culture); ", "; _
                     DateTimeKindHelper.ToString(localDate.Kind)
        Debug.Print "   UTC date and time: "; utcDate.ToString3(culture); ", "; _
                     DateTimeKindHelper.ToString(utcDate.Kind)
    Next
End Sub

' The example displays the following output:
'       English (United States):
'          Local date and time: 6/19/2015 10:35:50 AM, Local
'          UTC date and time: 6/19/2015 5:35:50 PM, Utc
'
'       English (United Kingdom):
'          Local date and time: 19/06/2015 10:35:50, Local
'          UTC date and time: 19/06/2015 17:35:50, Utc
'
'       français (France):
'          Local date and time: 19/06/2015 10:35:50, Local
'          UTC date and time: 19/06/2015 17:35:50, Utc
'
'       Deutsch (Deutschland):
'          Local date and time: 19.06.2015 10:35:50, Local
'          UTC date and time: 19.06.2015 17:35:50, Utc
'
'       русский (Россия):
'          Local date and time: 19.06.2015 10:35:50, Local
'          UTC date and time: 19.06.2015 17:35:50, Utc

